"""Run each mode in a fresh Linux process; times exclude input construction."""
import argparse
import inspect
import json
from pathlib import Path
import resource
import tempfile
from time import perf_counter

import excelsior

parser = argparse.ArgumentParser()
parser.add_argument('--mode', choices=['editor', 'stream', 'copy'], default='editor')
parser.add_argument('--rows', type=int, default=100_000)
parser.add_argument('--sheets', type=int, default=1)
parser.add_argument('--compression-level', type=int, default=3)
parser.add_argument('--generator', action='store_true')
parser.add_argument('--label', default='current')
parser.add_argument('--verify', action='store_true')
parser.add_argument('--temp-dir', type=Path, help='Use a disk-backed directory for XLSX/spool files')
args = parser.parse_args()

if args.generator:
    class Frame:
        height = args.rows
        columns = ['id', 'name', *[f'n{i}' for i in range(6)]]
        dtypes = ['String', 'String', *['Int64'] * 6]
        def iter_rows(self):
            for number in range(self.height):
                yield f'{number:020d}', f'row {number}', *([number] * 6)
    frame = Frame()
else:
    import polars as pl
    frame = pl.DataFrame({
        'id': [f'{i:020d}' for i in range(args.rows)],
        'name': [f'row {i}' for i in range(args.rows)],
        **{f'n{column}': list(range(args.rows)) for column in range(6)},
    })

if args.temp_dir is not None:
    args.temp_dir.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(prefix='excelsior-bench-', dir=args.temp_dir) as directory:
    output = Path(directory) / 'result.xlsx'
    if args.mode == 'stream':
        writer = excelsior.WorkbookWriter(output, compression_level=args.compression_level)
        compression = args.compression_level
    else:
        writer = excelsior.Editor.create(output)
        compression = args.compression_level if 'compression_level' in inspect.signature(writer.save).parameters else 'legacy-1-with-stored-small-parts'
    started = perf_counter()
    for sheet in range(args.sheets):
        if sheet:
            if args.mode == 'stream':
                writer.add_worksheet(f'Sheet{sheet + 1}')
            else:
                writer = writer.add_worksheet(f'Sheet{sheet + 1}')
        writer.with_polars(frame)
    written = perf_counter()
    if args.mode == 'stream' or isinstance(compression, str):
        writer.save()
    else:
        writer.save(compression_level=args.compression_level)
    saved = perf_counter()
    copies = None
    if args.mode == 'copy':
        begin = perf_counter()
        for _ in range(5):
            branch = writer.copy()
        copies = perf_counter() - begin
    result = {
        'label': args.label, 'version': excelsior.__version__, 'package': excelsior.__file__,
        'mode': args.mode, 'rows_per_sheet': args.rows, 'columns': 8, 'sheets': args.sheets,
        'generator': args.generator, 'compression_level': compression,
        'write_seconds': written - started, 'save_seconds': saved - written,
        'total_seconds': saved - started, 'five_copies_seconds': copies,
        'peak_rss_mib': resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024,
        'xlsx_bytes': output.stat().st_size,
        'temporary_files_after_save': len(list(Path(directory).glob('.excelsior-*'))),
    }
    if args.verify:
        from zipfile import ZipFile
        from xml.parsers import expat
        with ZipFile(output) as archive:
            for sheet in range(1, args.sheets + 1):
                state = {'rows': 0, 'row': 0, 'cell': None, 'kind': None, 'capture': False, 'text': [], 'last': {}}
                def start(name, attributes):
                    name = name.rsplit('}', 1)[-1]
                    if name == 'row':
                        state['rows'] += 1
                        state['row'] = int(attributes['r'])
                    elif name == 'c':
                        state['cell'] = attributes['r']
                        state['kind'] = attributes.get('t', 'n')
                        state['text'] = []
                    elif name in ('t', 'v') and state['row'] == args.rows + 1:
                        state['capture'] = True
                def text(value):
                    if state['capture']:
                        state['text'].append(value)
                def end(name):
                    name = name.rsplit('}', 1)[-1]
                    if name in ('t', 'v'):
                        state['capture'] = False
                    elif name == 'c' and state['row'] == args.rows + 1:
                        state['last'][state['cell']] = (state['kind'], ''.join(state['text']))
                validator = expat.ParserCreate(namespace_separator='}')
                validator.StartElementHandler = start
                validator.CharacterDataHandler = text
                validator.EndElementHandler = end
                with archive.open(f'xl/worksheets/sheet{sheet}.xml') as source:
                    while block := source.read(256 * 1024):
                        validator.Parse(block, False)
                    validator.Parse(b'', True)
                last = args.rows + 1
                assert state['rows'] == last
                assert state['last'][f'A{last}'] == ('inlineStr', f'{args.rows - 1:020d}')
                assert state['last'][f'H{last}'] == ('n', str(args.rows - 1))
        result['verified_rows_and_last_values'] = True
    print(json.dumps(result))
