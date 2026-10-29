from copy import copy, deepcopy
from datetime import date, datetime
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_STORED, ZIP_DEFLATED
import xml.etree.ElementTree as ET

import openpyxl
import polars as pl
import pytest

from excelsior import Editor, Scanner, WorkbookWriter
from helpers_excel import assert_with_libreoffice

SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
DOC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'


def read_book(path):
    return openpyxl.load_workbook(path)


def rewrite_zip(path, changes):
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts.update(changes)
    with ZipFile(path, 'w') as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


def test_editor_branches_keep_original_archive_after_another_branch_saves(tmp_path):
    path, snapshot = tmp_path / 'original.xlsx', tmp_path / 'snapshot.xlsx'
    initial = Editor.create(path)
    initial.set_cell('A1', 'base')
    initial = initial.add_worksheet('Other')
    initial.set_cell('A1', 'other original')
    initial.save()
    original = Editor.open(path)
    branch = original.with_worksheet('Other')
    branch.set_cell('A1', 'branch')
    branch = branch.rename_worksheet('Sheet1', 'Renamed')
    branch.save()
    original.save(snapshot)
    book = read_book(snapshot)
    assert book.sheetnames == ['Sheet1', 'Other']
    assert book['Other']['A1'].value == 'other original'
    assert read_book(path)['Other']['A1'].value == 'branch'
    selected = original.with_worksheet('Sheet1')
    selected.set_cell('B1', 'independent')
    selected.save(tmp_path / 'selected.xlsx')
    original.save(snapshot)
    assert read_book(snapshot)['Sheet1']['B1'].value is None


def test_editor_retains_source_when_path_is_removed(tmp_path):
    path = tmp_path / 'removed.xlsx'
    editor = Editor.create(path)
    editor.set_cell('A1', 'retained')
    path.unlink()
    editor.save()
    assert read_book(path)['Sheet1']['A1'].value == 'retained'


@pytest.mark.parametrize('level', [0, 1, 3, 9])
def test_editor_compression_level_applies_to_every_zip_entry(tmp_path, level):
    path = tmp_path / f'compression-{level}.xlsx'
    editor = Editor.create(path)
    editor.with_polars(pl.DataFrame({'id': ['00123'] * 100, 'n': range(100)}))
    editor.save(compression_level=level)
    with ZipFile(path) as archive:
        assert {info.compress_type for info in archive.infolist()} == {ZIP_STORED if level == 0 else ZIP_DEFLATED}
        assert archive.testzip() is None
    assert read_book(path)['Sheet1']['A2'].value == '00123'


def test_default_compression_and_invalid_levels_preserve_destination(tmp_path):
    path = tmp_path / 'default.xlsx'
    editor = Editor.create(path)
    editor.set_cell('A1', 'value')
    editor.save()
    original = path.read_bytes()
    for level in [-1, 10]:
        with pytest.raises(ValueError, match='compression_level'):
            editor.save(compression_level=level)
        assert path.read_bytes() == original
        with pytest.raises(ValueError, match='compression_level'):
            WorkbookWriter(path, compression_level=level)
        assert path.read_bytes() == original
    with ZipFile(path) as archive:
        assert all(info.compress_type == ZIP_DEFLATED for info in archive.infolist())


@pytest.mark.parametrize('level', [0, 3, 9])
def test_streaming_dataframe_multiple_sheets_types_formats_and_dimensions(tmp_path, level):
    path = tmp_path / 'streaming.xlsx'
    frame = pl.DataFrame({'id': ['00123', '=1+1'], 'n': [1.5, 2.5], 'flag': [True, False], 'date': [date(2025, 1, 2)] * 2})
    with WorkbookWriter(path, 'Data & "values"', compression_level=level) as writer:
        writer.set_column_width('B:C', 24)
        assert writer.with_polars(frame, 'B3') is None
        assert writer.add_worksheet('Second') is None
        writer.with_polars(frame.slice(1))
    book = read_book(path)
    ws = book['Data & "values"']
    assert ws['B4'].value == '00123' and ws['B4'].data_type == 's'
    assert ws['B5'].value == '=1+1' and ws['B5'].data_type == 's'
    assert ws['C4'].value == 1.5 and ws['C4'].number_format == '0.###############'
    assert ws['D4'].value is True
    assert ws['E4'].value == datetime(2025, 1, 2)
    assert ws.column_dimensions['B'].number_format == '@'
    assert ws.column_dimensions['B'].width == 24
    assert book['Second']['A2'].value == '=1+1'
    with ZipFile(path) as archive:
        assert archive.testzip() is None
        assert {info.compress_type for info in archive.infolist()} == {ZIP_STORED if level == 0 else ZIP_DEFLATED}
    readonly = openpyxl.load_workbook(path, read_only=True)
    assert readonly['Data & "values"'].max_row == 5
    assert readonly['Data & "values"'].max_column == 5
    readonly.close()
    if level == 3:
        assert_with_libreoffice(path)


def test_streaming_rows_are_typed_and_column_formats_can_be_set_after_rows(tmp_path):
    path = tmp_path / 'rows.xlsx'
    writer = WorkbookWriter(path)
    writer.write_rows((row for row in [('00123', 1.25, True), ('=1+1', 2.5, False)]))
    writer.set_number_format('B:', '0.000')
    writer.save()
    ws = read_book(path)['Sheet1']
    assert ws['A1'].value == '00123'
    assert ws['A2'].value == '=1+1' and ws['A2'].data_type == 's'
    assert ws['B1'].number_format == ws['B2'].number_format == '0.000'
    assert ws['C2'].value is False
    with pytest.raises(RuntimeError, match='finalized'):
        writer.append_row([3])
    with pytest.raises(TypeError):
        copy(writer)
    with pytest.raises(TypeError):
        deepcopy(writer)


def test_streaming_iterator_failure_rolls_back_current_call(tmp_path):
    path = tmp_path / 'rollback.xlsx'
    writer = WorkbookWriter(path)
    writer.append_row(['before'])
    def broken():
        yield ['discard']
        raise ValueError('broken input')
    with pytest.raises(ValueError, match='broken input'):
        writer.write_rows(broken())
    writer.append_row(['after'])
    writer.save()
    assert [row[0] for row in read_book(path)['Sheet1'].values] == ['before', 'after']


def test_streaming_empty_frames_and_nonfinite_values(tmp_path):
    path = tmp_path / 'empty.xlsx'
    with WorkbookWriter(path) as writer:
        writer.with_polars(pl.DataFrame(schema={'id': pl.String, 'n': pl.Float64}))
        writer.add_worksheet('Values')
        writer.with_polars(pl.DataFrame({'n': [float('nan'), float('inf')]}))
    book = read_book(path)
    assert list(book['Sheet1'].values) == [('id', 'n')]
    assert book['Sheet1'].column_dimensions['A'].number_format == '@'
    assert book['Values']['A2'].value == 'nan'
    assert book['Values']['A3'].value == 'inf'


def test_streaming_validation_and_abort_preserve_previous_file(tmp_path):
    path = tmp_path / 'previous.xlsx'
    path.write_bytes(b'preserve')
    with pytest.raises(ValueError, match='stop'):
        with WorkbookWriter(path) as writer:
            writer.append_row(['not committed'])
            raise ValueError('stop')
    assert path.read_bytes() == b'preserve'
    assert not list(tmp_path.glob('.excelsior-*'))
    writer = WorkbookWriter(path)
    writer.append_row(['before'])
    with pytest.raises(RuntimeError, match='overlapping'):
        writer.write_rows([['bad']], 'A1')
    with pytest.raises(RuntimeError):
        writer.add_worksheet('Sheet1')
    writer.append_row(['after'])
    writer.save()
    assert list(read_book(path)['Sheet1'].values) == [('before',), ('after',)]


def test_streaming_save_as_retains_original_destination(tmp_path):
    source, copy_path = tmp_path / 'source.xlsx', tmp_path / 'copy.xlsx'
    source.write_bytes(b'previous')
    writer = WorkbookWriter(source)
    writer.append_row(['00123'])
    writer.save(copy_path)
    assert source.read_bytes() == b'previous'
    writer.save()
    assert source.read_bytes() == copy_path.read_bytes()


def test_streaming_does_not_import_or_convert_through_arrow(tmp_path):
    class Frame:
        columns = ['id']
        def iter_rows(self): yield ('00123',)
        def to_arrow(self): raise AssertionError('Arrow must not be used')
    path = tmp_path / 'protocol.xlsx'
    writer = WorkbookWriter(path)
    writer.with_polars(Frame())
    writer.save()
    assert read_book(path)['Sheet1']['A2'].value == '00123'


def make_prefixed_parts(path):
    for prefix, uri in [('m', SHEET_NS), ('p', REL_NS), ('r', DOC_REL_NS), ('ct', CT_NS)]:
        ET.register_namespace(prefix, uri)
    with ZipFile(path) as archive:
        parts = {}
        for name in archive.namelist():
            content = archive.read(name)
            if name.endswith('.xml') or name.endswith('.rels'):
                content = ET.tostring(ET.fromstring(content), encoding='utf-8', xml_declaration=True)
                content = re.sub(rb'([\w:]+)="([^"]*)"', lambda m: m[1] + b" = '" + m[2].replace(b"'", b'&apos;') + b"'", content)
            parts[name] = content
    rewrite_zip(path, parts)


def test_prefixed_single_quoted_metadata_styles_rows_and_unknown_extensions(tmp_path):
    path = tmp_path / 'prefixes.xlsx'
    Editor.create(path, 'Data')
    make_prefixed_parts(path)
    sheet = f'''<?xml version='1.0' encoding='UTF-8'?>
    <m:worksheet xmlns:m='{SHEET_NS}' xmlns:x='urn:extension' x:note='a > "quoted" &amp; b'>
      <!-- <row r="999"><c r="XFD999"/></row> -->
      <m:dimension ref='A1:B3'/><m:sheetData>
      <m:row r = '1'><m:c t='inlineStr' r='A1' s='0'><m:is><m:t><![CDATA[literal <row r="999"> & data]]></m:t></m:is></m:c><m:c r='B1'><m:v>9</m:v></m:c></m:row>
      <m:row r='3' />
      </m:sheetData><m:extLst><x:row r='555'><x:c r='XFD555'/></x:row></m:extLst>
    </m:worksheet>'''.encode()
    rewrite_zip(path, {'xl/worksheets/sheet1.xml': sheet})
    editor = Editor.open(path)
    editor = editor.set_font('A1', 'Font & "Name"', 12).set_fill('A1', 'FF0000')
    editor.with_polars(pl.DataFrame({'id': ['00123'], 'n': [42]}), 'B2')
    editor = editor.set_number_format('C:', '0.00" & unit"')
    assert editor.last_rows_index('A:C') == [1, 3, 3]
    editor = editor.add_worksheet('Other').with_worksheet('Data').rename_worksheet('Data', 'Renamed & data')
    editor = editor.merge_cells('D1:E1')
    editor.save()
    book = read_book(path)
    ws = book['Renamed & data']
    assert ws['A1'].value == 'literal <row r="999"> & data'
    assert ws['A1'].font.name == 'Font & "Name"'
    assert ws['B3'].value == '00123'
    assert ws['C3'].value == 42
    assert ws['C3'].number_format == '0.00" & unit"'
    assert Scanner(path).get_sheets() == ['Renamed & data', 'Other']
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        assert root.attrib['{urn:extension}note'] == 'a > "quoted" & b'
        assert root.find(f'{{{SHEET_NS}}}extLst/{{urn:extension}}row') is not None
    readonly = openpyxl.load_workbook(path, read_only=True)
    assert readonly['Renamed & data'].max_row == 3
    readonly.close()


def test_different_default_namespace_is_preserved(tmp_path):
    path = tmp_path / 'default-other.xlsx'
    Editor.create(path)
    sheet = f'''<m:worksheet xmlns:m='{SHEET_NS}' xmlns='urn:other'><m:sheetData><m:row r='1'><m:c r='A1' t='inlineStr'><m:is><m:t>original</m:t></m:is></m:c></m:row></m:sheetData><m:extLst><unknown foo='kept'/></m:extLst></m:worksheet>'''.encode()
    rewrite_zip(path, {'xl/worksheets/sheet1.xml': sheet})
    editor = Editor.open(path)
    editor.set_cell('B1', 'new')
    editor.append_row(['next'])
    editor = editor.set_fill('C3', 'FF0000')
    editor.save()
    ws = read_book(path)['Sheet1']
    assert ws['A1'].value == 'original'
    assert ws['B1'].value == 'new'
    assert ws['A2'].value == 'next'
    assert ws['C3'].has_style
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        assert root.find(f'{{{SHEET_NS}}}extLst/{{urn:other}}unknown') is not None


def test_malformed_xml_raises_normal_error(tmp_path):
    path = tmp_path / 'malformed.xlsx'
    Editor.create(path)
    rewrite_zip(path, {'xl/worksheets/sheet1.xml': b'<worksheet><sheetData><row></sheetData></worksheet>'})
    with pytest.raises(RuntimeError):
        Editor.open(path)


def test_builtin_date_format_survives_font_and_fill_changes(tmp_path):
    path = tmp_path / 'builtin-date.xlsx'
    book = openpyxl.Workbook()
    book.active['A1'] = datetime(2025, 1, 2)
    book.active['A1'].number_format = 'mm-dd-yy'
    book.save(path)
    editor = Editor.open(path)
    editor = editor.set_font('A1', 'Arial', 12).set_fill('A1', 'FF0000')
    editor.save()
    cell = read_book(path).active['A1']
    assert cell.value == datetime(2025, 1, 2)
    assert cell.number_format == 'mm-dd-yy'


def test_column_outline_and_custom_attributes_survive_width_and_format_changes(tmp_path):
    path = tmp_path / 'outline.xlsx'
    book = openpyxl.Workbook()
    book.active['A1'] = 123
    book.active.column_dimensions['A'].outlineLevel = 2
    book.active.column_dimensions['A'].collapsed = True
    book.save(path)
    editor = Editor.open(path).set_column_width('A', 24).set_number_format('A:', '0.00')
    editor.save()
    column = read_book(path).active.column_dimensions['A']
    assert column.outlineLevel == 2 and column.collapsed
    assert column.width == 24


def test_parallel_saves_use_independent_source_zip_cursors(tmp_path):
    from concurrent.futures import ThreadPoolExecutor
    path = tmp_path / 'parallel-source.xlsx'
    Editor.create(path)
    payload = bytes(range(256)) * 4096
    rewrite_zip(path, {'xl/media/unreferenced.bin': payload})
    editor = Editor.open(path)
    branches = [editor.copy() for _ in range(6)]
    outputs = [tmp_path / f'parallel-{index}.xlsx' for index in range(6)]
    with ThreadPoolExecutor(max_workers=6) as pool:
        list(pool.map(lambda item: item[0].save(item[1], compression_level=0), zip(branches, outputs)))
    for output in outputs:
        with ZipFile(output) as archive:
            assert archive.testzip() is None
            assert archive.read('xl/media/unreferenced.bin') == payload


def test_successful_streaming_save_cleans_up_spools_even_with_writer_alive(tmp_path):
    path = tmp_path / 'clean.xlsx'
    writer = WorkbookWriter(path)
    writer.append_row(['before'])
    assert list(tmp_path.glob('.excelsior-*'))
    writer.save()
    assert not list(tmp_path.glob('.excelsior-*'))
    writer.save(tmp_path / 'copy.xlsx')
    assert not list(tmp_path.glob('.excelsior-*'))
    assert read_book(path)['Sheet1']['A1'].value == 'before'


def test_utf8_bom_is_preserved_and_does_not_change_element_offsets(tmp_path):
    path = tmp_path / 'bom.xlsx'
    Editor.create(path)
    with ZipFile(path) as archive:
        parts = {name: b'\xef\xbb\xbf' + archive.read(name) for name in archive.namelist()}
    rewrite_zip(path, parts)
    editor = Editor.open(path)
    editor.with_polars(pl.DataFrame({'id': ['00123']}))
    editor = editor.set_fill('A1', 'FF0000').add_worksheet('Other')
    editor.append_row(['value'])
    editor.save()
    book = read_book(path)
    assert book['Sheet1']['A2'].value == '00123'
    assert book['Other']['A1'].value == 'value'


def test_dimensions_of_cow_branches_do_not_leak(tmp_path):
    path = tmp_path / 'dimensions-cow.xlsx'
    editor = Editor.create(path)
    editor.with_polars(pl.DataFrame({'id': ['00123']}))
    branch = editor.copy()
    branch.set_cell('Z100', 'other')
    branch.save(tmp_path / 'branch.xlsx')
    editor.save()
    before = openpyxl.load_workbook(path, read_only=True)
    after = openpyxl.load_workbook(tmp_path / 'branch.xlsx', read_only=True)
    assert (before.active.max_row, before.active.max_column) == (2, 1)
    assert (after.active.max_row, after.active.max_column) == (100, 26)
    before.close()
    after.close()


def test_indexed_scalar_edits_survive_cache_invalidation_and_cow(tmp_path):
    import random
    path = tmp_path / 'indexed.xlsx'
    editor = Editor.create(path)
    editor.append_table_at([[f'original-{row}', 'payload'] for row in range(1, 41)], 'A1')
    editor.set_cell('A1', 'indexed')  # construct the lazy row index
    snapshot = editor.copy()
    expected = {f'A{row}': f'original-{row}' for row in range(1, 41)}
    expected['A1'] = 'indexed'
    rng = random.Random(31415)
    for change in range(100):
        coord = f'{rng.choice(["A", "B", "Z"])}{rng.randint(1, 40)}'
        value = f'value-{change}-' + 'long' * rng.randrange(12)
        editor.set_cell(coord, value)
        expected[coord] = value
        if change == 20:
            editor = editor.set_column_width('B', 24)
        elif change == 40:
            editor = editor.set_fill('A3', 'FF0000')
        elif change == 60:
            editor = editor.merge_cells('H1:I1')
        elif change == 80:
            editor.append_row(['tail'])
    editor.set_cell('A50', 'inserted')
    expected['A50'] = 'inserted'
    editor.set_cell('B45', 'sparse')
    expected['B45'] = 'sparse'
    editor.set_cell('A1', 'short')
    expected['A1'] = 'short'
    editor.save()
    snapshot.set_cell('B2', 'branch-only')
    copy_path = tmp_path / 'snapshot.xlsx'
    snapshot.save(copy_path)
    sheet = read_book(path).active
    for coord, value in expected.items():
        assert sheet[coord].value == value
    before = read_book(copy_path).active
    assert before['A1'].value == 'indexed'
    assert before['B2'].value == 'branch-only'
    assert before['A50'].value is None
    editor = editor.add_worksheet('Other')
    editor.append_row(['other'])
    editor = editor.with_worksheet('Sheet1')
    editor.set_cell('A1', 'after-switch')
    editor.save()
    assert read_book(path)['Sheet1']['A1'].value == 'after-switch'


def test_indexed_rows_inherit_namespaces_declared_on_sheet_data(tmp_path):
    path = tmp_path / 'row-scope.xlsx'
    Editor.create(path)
    sheet = f'''<m:worksheet xmlns:m='{SHEET_NS}'><m:sheetData xmlns:q='{SHEET_NS}'>
    <q:row r='1' ht='23'><q:c r='A1' t='inlineStr'><q:is><q:t>before</q:t></q:is></q:c></q:row>
    <q:row r='2'><q:c r='A2'><q:v>12</q:v></q:c></q:row>
    </m:sheetData></m:worksheet>'''.encode()
    rewrite_zip(path, {'xl/worksheets/sheet1.xml': sheet})
    editor = Editor.open(path)
    for text in ['a much longer value', 'x', 'final']:
        editor.set_cell('A1', text)
        editor.set_cell('B2', text)
    editor.save()
    ws = read_book(path).active
    assert ws['A1'].value == 'final' and ws['A2'].value == 12
    assert ws['B2'].value == 'final'
    assert ws.row_dimensions[1].height == 23
