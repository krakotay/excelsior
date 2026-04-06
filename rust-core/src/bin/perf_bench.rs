use std::env;
use std::cell::RefCell;
use std::fs;
use std::path::{Path, PathBuf};
use std::rc::Rc;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use anyhow::{Context, Result, bail};
use rust_core::style::{AlignSpec, FontSpec, HorizAlignment, StyleSpec, VertAlignment};
use rust_core::{XlsxEditor, scan};

const EXCEL_MAX_ROWS: usize = 1_048_576;

#[derive(Clone)]
struct ScenarioConfig {
    preset_name: String,
    append_rows: usize,
    append_cols: usize,
    overlap_seed_rows: usize,
    overlap_patch_rows: usize,
    style_seed_rows: usize,
    style_cols: usize,
    style_waves: usize,
    switches: usize,
}

impl ScenarioConfig {
    fn quick() -> Self {
        Self {
            preset_name: "quick".to_string(),
            append_rows: 5_000,
            append_cols: 16,
            overlap_seed_rows: 4_000,
            overlap_patch_rows: 4_000,
            style_seed_rows: 8_000,
            style_cols: 16,
            style_waves: 8,
            switches: 40,
        }
    }

    fn full() -> Self {
        Self {
            preset_name: "full".to_string(),
            append_rows: 75_000,
            append_cols: 20,
            overlap_seed_rows: 40_000,
            overlap_patch_rows: 40_000,
            style_seed_rows: 50_000,
            style_cols: 20,
            style_waves: 12,
            switches: 150,
        }
    }

    fn million() -> Self {
        Self {
            preset_name: "million".to_string(),
            append_rows: 1_000_000,
            append_cols: 10,
            overlap_seed_rows: 180_000,
            overlap_patch_rows: 180_000,
            style_seed_rows: 150_000,
            style_cols: 24,
            style_waves: 18,
            switches: 400,
        }
    }

    fn hardcore() -> Self {
        Self {
            preset_name: "hardcore".to_string(),
            append_rows: 250_000,
            append_cols: 32,
            overlap_seed_rows: 120_000,
            overlap_patch_rows: 120_000,
            style_seed_rows: 180_000,
            style_cols: 32,
            style_waves: 24,
            switches: 500,
        }
    }

    fn validate(&self) -> Result<()> {
        if self.append_rows == 0 || self.append_cols == 0 {
            bail!("append benchmark needs rows > 0 and cols > 0");
        }
        if self.append_rows > EXCEL_MAX_ROWS {
            bail!("append_rows exceeds Excel row limit of {EXCEL_MAX_ROWS}");
        }
        if self.overlap_seed_rows == 0 || self.overlap_patch_rows == 0 {
            bail!("overlap benchmark needs positive seed and patch row counts");
        }
        if self.overlap_seed_rows > EXCEL_MAX_ROWS || self.overlap_patch_rows > EXCEL_MAX_ROWS {
            bail!("overlap rows exceed Excel row limit of {EXCEL_MAX_ROWS}");
        }
        if self.style_seed_rows == 0 || self.style_cols == 0 {
            bail!("style benchmark needs positive style rows and cols");
        }
        if self.style_seed_rows > EXCEL_MAX_ROWS {
            bail!("style_seed_rows exceeds Excel row limit of {EXCEL_MAX_ROWS}");
        }
        Ok(())
    }
}

struct Cli {
    config: ScenarioConfig,
    output_dir: PathBuf,
}

struct BenchResult {
    scenario: &'static str,
    open_ms: f64,
    mutate_ms: f64,
    save_ms: f64,
    total_ms: f64,
    output_bytes: u64,
    input_path: PathBuf,
    output_path: PathBuf,
    details: Vec<BenchDetail>,
}

#[derive(Debug)]
struct BenchDetail {
    label: &'static str,
    elapsed_ms: f64,
}

struct GeneratedRows {
    next_row: usize,
    total_rows: usize,
    cols: usize,
    salt: u64,
}

impl GeneratedRows {
    fn new(total_rows: usize, cols: usize, salt: u64) -> Self {
        Self {
            next_row: 0,
            total_rows,
            cols,
            salt,
        }
    }
}

impl Iterator for GeneratedRows {
    type Item = Vec<String>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.next_row >= self.total_rows {
            return None;
        }

        let row_idx = self.next_row;
        self.next_row += 1;

        let mut row = Vec::with_capacity(self.cols);
        for col_idx in 0..self.cols {
            let mix = (row_idx as u64)
                .wrapping_mul(31)
                .wrapping_add((col_idx as u64).wrapping_mul(17))
                .wrapping_add(self.salt);
            row.push(format!("r{row_idx}_c{col_idx}_v{}", mix % 10_007));
        }
        Some(row)
    }
}

fn main() -> Result<()> {
    let cli = parse_cli()?;
    cli.config.validate()?;

    let root = workspace_root()?;
    let test_dir = root.join("test");

    fs::create_dir_all(&cli.output_dir)
        .with_context(|| format!("failed to create {}", cli.output_dir.display()))?;

    println!("bench go brrr");
    println!("output dir: {}", cli.output_dir.display());

    let results = vec![
        bench_append_tail(&test_dir, &cli.output_dir, &cli.config)?,
        bench_append_overlap(&test_dir, &cli.output_dir, &cli.config)?,
        bench_style_storm(&test_dir, &cli.output_dir, &cli.config)?,
        bench_style_precision(&test_dir, &cli.output_dir, &cli.config)?,
        bench_format_matrix(&test_dir, &cli.output_dir, &cli.config)?,
        bench_switch_storm(&test_dir, &cli.output_dir, &cli.config)?,
    ];

    print_report(&results, &cli.config);
    write_markdown_report(&cli.output_dir, &results, &cli.config)?;
    Ok(())
}

fn parse_cli() -> Result<Cli> {
    let root = workspace_root()?;
    let mut config = ScenarioConfig::quick();
    let mut output_dir: Option<PathBuf> = None;

    let args: Vec<String> = env::args().skip(1).collect();
    let mut i = 0usize;
    while i < args.len() {
        match args[i].as_str() {
            "--quick" => config = ScenarioConfig::quick(),
            "--full" => config = ScenarioConfig::full(),
            "--million" => config = ScenarioConfig::million(),
            "--hardcore" => config = ScenarioConfig::hardcore(),
            "--rows" => {
                i += 1;
                config.append_rows = parse_usize_arg(&args, i, "--rows")?;
            }
            "--cols" => {
                i += 1;
                config.append_cols = parse_usize_arg(&args, i, "--cols")?;
            }
            "--overlap-seed-rows" => {
                i += 1;
                config.overlap_seed_rows = parse_usize_arg(&args, i, "--overlap-seed-rows")?;
            }
            "--overlap-patch-rows" => {
                i += 1;
                config.overlap_patch_rows = parse_usize_arg(&args, i, "--overlap-patch-rows")?;
            }
            "--style-rows" => {
                i += 1;
                config.style_seed_rows = parse_usize_arg(&args, i, "--style-rows")?;
            }
            "--style-cols" => {
                i += 1;
                config.style_cols = parse_usize_arg(&args, i, "--style-cols")?;
            }
            "--style-waves" => {
                i += 1;
                config.style_waves = parse_usize_arg(&args, i, "--style-waves")?;
            }
            "--switches" => {
                i += 1;
                config.switches = parse_usize_arg(&args, i, "--switches")?;
            }
            "--output-dir" => {
                i += 1;
                let raw = args
                    .get(i)
                    .with_context(|| "missing value for --output-dir".to_string())?;
                output_dir = Some(PathBuf::from(raw));
            }
            other => bail!("unknown argument: {other}"),
        }
        i += 1;
    }

    let output_dir = output_dir.unwrap_or_else(|| {
        root.join("test")
            .join("perf_outputs")
            .join(format!("{}-{}", config.preset_name, unix_timestamp()))
    });

    Ok(Cli { config, output_dir })
}

fn parse_usize_arg(args: &[String], idx: usize, flag: &str) -> Result<usize> {
    args.get(idx)
        .with_context(|| format!("missing value for {flag}"))?
        .parse::<usize>()
        .with_context(|| format!("failed to parse numeric value for {flag}"))
}

fn bench_append_tail(test_dir: &Path, out_dir: &Path, cfg: &ScenarioConfig) -> Result<BenchResult> {
    println!("run append_tail");
    let src = staged_copy(test_dir, out_dir, "test_new_ws.xlsx", "append_tail_input.xlsx")?;
    let dst = out_dir.join("append_tail_out.xlsx");
    run_bench("append_tail", &src, &dst, move |editor| {
        editor.append_table(GeneratedRows::new(cfg.append_rows, cfg.append_cols, 11))?;
        Ok(())
    })
}

fn bench_append_overlap(
    test_dir: &Path,
    out_dir: &Path,
    cfg: &ScenarioConfig,
) -> Result<BenchResult> {
    println!("prepare append_overlap input");
    let seed = out_dir.join("append_overlap_input.xlsx");
    prepare_large_sheet_input(
        test_dir,
        &seed,
        cfg.overlap_seed_rows,
        cfg.append_cols,
        "seed overlap input",
        101,
    )?;

    let dst = out_dir.join("append_overlap_out.xlsx");
    println!("run append_overlap");
    run_bench("append_overlap", &seed, &dst, move |editor| {
        editor.append_table_at(
            "A1",
            GeneratedRows::new(cfg.overlap_patch_rows, cfg.append_cols, 202),
        )?;
        Ok(())
    })
}

fn bench_style_storm(
    test_dir: &Path,
    out_dir: &Path,
    cfg: &ScenarioConfig,
) -> Result<BenchResult> {
    println!("prepare style_storm input");
    let src = out_dir.join("style_storm_input.xlsx");
    prepare_style_input(test_dir, &src, cfg)?;

    let dst = out_dir.join("style_storm_out.xlsx");
    println!("run style_storm");
    let details = Rc::new(RefCell::new(Vec::new()));
    let details_for_run = Rc::clone(&details);
    let result = run_bench("style_storm", &src, &dst, move |editor| {
        let wave_height = (cfg.style_seed_rows / cfg.style_waves.max(1)).max(1);
        let colors = [
            "FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC", "EAD1DC", "CFE2F3", "FFD966", "B6D7A8",
        ];
        let fonts = ["Arial", "Calibri", "Tahoma", "Verdana"];
        let alignments = [
            AlignSpec {
                horiz: Some(HorizAlignment::Left),
                vert: Some(VertAlignment::Top),
                wrap: false,
            },
            AlignSpec {
                horiz: Some(HorizAlignment::Center),
                vert: Some(VertAlignment::Center),
                wrap: true,
            },
            AlignSpec {
                horiz: Some(HorizAlignment::Right),
                vert: Some(VertAlignment::Bottom),
                wrap: false,
            },
            AlignSpec {
                horiz: Some(HorizAlignment::Justify),
                vert: Some(VertAlignment::Justify),
                wrap: true,
            },
        ];

        for wave in 0..cfg.style_waves {
            let row_start = (wave * wave_height) as u32 + 1;
            let row_end = ((wave + 1) * wave_height).min(cfg.style_seed_rows) as u32;
            if row_start > row_end {
                break;
            }

            let col_cap = cfg.style_cols.saturating_sub((wave % 3) * 2).max(6);
            let col_end = excel_col_name((col_cap - 1) as u32);
            let rect = format!("A{row_start}:{col_end}{row_end}");
            let mid_col = excel_col_name(((col_cap / 2).saturating_sub(1)) as u32);
            let tail_col = excel_col_name((col_cap.saturating_sub(1)) as u32);
            let right_rect = format!("{mid_col}{row_start}:{tail_col}{row_end}");
            let color = colors[wave % colors.len()];
            let font = fonts[wave % fonts.len()];
            let align = &alignments[wave % alignments.len()];
            let rect_style = StyleSpec {
                fill_rgb: Some(color.to_string()),
                border_style: Some("thin".to_string()),
                align: Some(align.clone()),
                ..Default::default()
            };
            let font_style = StyleSpec {
                font: Some(FontSpec {
                    name: font.to_string(),
                    size: 10.0 + (wave % 4) as f32,
                    bold: wave % 2 == 0,
                    italic: wave % 3 == 0,
                }),
                ..Default::default()
            };

            let started = Instant::now();
            {
                editor
                    .set_style(&rect, &rect_style)?
                    .set_style(&right_rect, &font_style)?;
            }
            details_for_run.borrow_mut().push(BenchDetail {
                label: "wave_fill_border_font_align",
                elapsed_ms: as_millis(started.elapsed()),
            });
        }

        let started = Instant::now();
        {
            editor.with_worksheet("StyleAux")?;
            let aux_end = excel_col_name((cfg.style_cols.saturating_sub(1)) as u32);
            let aux_range = format!("A1:{aux_end}{}", (cfg.style_seed_rows / 4).max(1));
            editor.set_style(
                &aux_range,
                &StyleSpec {
                    fill_rgb: Some("CCE5FF".to_string()),
                    border_style: Some("thin".to_string()),
                    font: Some(FontSpec {
                        name: "Consolas".to_string(),
                        size: 11.0,
                        bold: true,
                        italic: false,
                    }),
                    align: Some(AlignSpec {
                        horiz: Some(HorizAlignment::Center),
                        vert: Some(VertAlignment::Center),
                        wrap: true,
                    }),
                    ..Default::default()
                },
            )?;
        }
        details_for_run.borrow_mut().push(BenchDetail {
            label: "style_aux",
            elapsed_ms: as_millis(started.elapsed()),
        });

        let started = Instant::now();
        {
            for col in 0..cfg.style_cols.min(12) {
                let col_name = excel_col_name(col as u32);
                let fmt = match col % 4 {
                    0 => "#,##0",
                    1 => "#,##0.00",
                    2 => "0.000%",
                    _ => "yyyy-mm-dd hh:mm:ss",
                };
                editor.set_number_format(&format!("{col_name}:"), fmt)?;
            }
        }
        details_for_run.borrow_mut().push(BenchDetail {
            label: "style_num_formats",
            elapsed_ms: as_millis(started.elapsed()),
        });
        Ok(())
    })?;
    Ok(BenchResult {
        details: Rc::try_unwrap(details).unwrap().into_inner(),
        ..result
    })
}

fn bench_style_precision(
    test_dir: &Path,
    out_dir: &Path,
    cfg: &ScenarioConfig,
) -> Result<BenchResult> {
    println!("prepare style_precision input");
    let src = out_dir.join("style_precision_input.xlsx");
    prepare_style_input(test_dir, &src, cfg)?;

    let dst = out_dir.join("style_precision_out.xlsx");
    println!("run style_precision");
    let details = Rc::new(RefCell::new(Vec::new()));
    let details_for_run = Rc::clone(&details);
    let result = run_bench("style_precision", &src, &dst, move |editor| {
        let rows = cfg.style_seed_rows.min(8_000).max(256) as u32;
        let cols = cfg.style_cols.min(20).max(8) as u32;
        let right = excel_col_name(cols - 1);
        let mid = excel_col_name((cols / 2).max(1) - 1);
        let bottom = rows;
        let stripe_colors = ["FFF2CC", "D9EAD3", "D0E0E3", "F4CCCC"];
        let fonts = ["Calibri", "Arial", "Verdana", "Tahoma"];
        let alignments = [
            AlignSpec {
                horiz: Some(HorizAlignment::Left),
                vert: Some(VertAlignment::Center),
                wrap: false,
            },
            AlignSpec {
                horiz: Some(HorizAlignment::Center),
                vert: Some(VertAlignment::Center),
                wrap: true,
            },
            AlignSpec {
                horiz: Some(HorizAlignment::Right),
                vert: Some(VertAlignment::Bottom),
                wrap: false,
            },
        ];

        let started = Instant::now();
        for layer in 0..12u32 {
            let row_start = 1 + layer * (rows / 12).max(1);
            let row_end = (row_start + (rows / 3)).min(bottom);
            let left = excel_col_name((layer % (cols / 2).max(1)) as u32);
            let range = format!("{left}{row_start}:{right}{row_end}");
            let focused = format!("{mid}{row_start}:{right}{row_end}");
            let fmt = match layer % 4 {
                0 => "#,##0",
                1 => "#,##0.00",
                2 => "0.000%",
                _ => "yyyy-mm-dd hh:mm:ss",
            };
            let base_style = StyleSpec {
                fill_rgb: Some(stripe_colors[layer as usize % stripe_colors.len()].to_string()),
                border_style: Some("thin".to_string()),
                align: Some(alignments[layer as usize % alignments.len()].clone()),
                ..Default::default()
            };
            let focused_style = StyleSpec {
                num_fmt_code: Some(fmt.to_string()),
                font: Some(FontSpec {
                    name: fonts[layer as usize % fonts.len()].to_string(),
                    size: 10.0 + (layer % 4) as f32,
                    bold: layer % 2 == 0,
                    italic: layer % 3 == 0,
                }),
                ..Default::default()
            };

            editor
                .set_style(&range, &base_style)?
                .set_style(&focused, &focused_style)?;
        }
        details_for_run.borrow_mut().push(BenchDetail {
            label: "style_layers",
            elapsed_ms: as_millis(started.elapsed()),
        });
        Ok(())
    })?;
    Ok(BenchResult {
        details: Rc::try_unwrap(details).unwrap().into_inner(),
        ..result
    })
}

fn bench_format_matrix(
    test_dir: &Path,
    out_dir: &Path,
    cfg: &ScenarioConfig,
) -> Result<BenchResult> {
    println!("prepare format_matrix input");
    let src = out_dir.join("format_matrix_input.xlsx");
    prepare_large_sheet_input(
        test_dir,
        &src,
        cfg.style_seed_rows,
        cfg.style_cols,
        "format input",
        505,
    )?;

    let dst = out_dir.join("format_matrix_out.xlsx");
    println!("run format_matrix");
    let details = Rc::new(RefCell::new(Vec::new()));
    let details_for_run = Rc::clone(&details);
    let result = run_bench("format_matrix", &src, &dst, move |editor| {
        let rows = cfg.style_seed_rows as u32;
        let cols = cfg.style_cols.max(8) as u32;
        let right = excel_col_name(cols.saturating_sub(1));
        let half_right = excel_col_name((cols / 2).max(1) - 1);
        let rect_top = rows.min((rows / 2).max(1));
        let rect_mid_start = (rows / 5).max(1);
        let rect_mid_end = (rect_mid_start + rows / 2).min(rows);
        let rect_bottom_start = (rows.saturating_mul(2) / 3).max(1);

        let started = Instant::now();
        editor.set_number_format(&format!("A1:{right}{rect_top}"), "#,##0.00")?;
        details_for_run.borrow_mut().push(BenchDetail {
            label: "rect_number",
            elapsed_ms: as_millis(started.elapsed()),
        });

        let started = Instant::now();
        editor.set_number_format(
            &format!("B{rect_mid_start}:{right}{rect_mid_end}"),
            "yyyy-mm-dd hh:mm:ss",
        )?;
        details_for_run.borrow_mut().push(BenchDetail {
            label: "rect_date",
            elapsed_ms: as_millis(started.elapsed()),
        });

        let started = Instant::now();
        editor.set_number_format(
            &format!("A{}:{}{}", (rows / 3).max(1), half_right, rows),
            "@",
        )?;
        details_for_run.borrow_mut().push(BenchDetail {
            label: "rect_text",
            elapsed_ms: as_millis(started.elapsed()),
        });

        let started = Instant::now();
        editor.set_number_format(
            &format!("{}{}:{right}{rows}", half_right, rect_bottom_start),
            r#"[Blue]#,##0.000;[Red]-#,##0.000;0.000;@"#,
        )?;
        details_for_run.borrow_mut().push(BenchDetail {
            label: "rect_custom",
            elapsed_ms: as_millis(started.elapsed()),
        });

        let row_targets = [
            2,
            (rows / 8).max(1),
            (rows / 4).max(1),
            (rows / 2).max(1),
            rows.saturating_sub(1).max(1),
        ];
        let row_formats = ["#,##0", "yyyy-mm-dd", "@", "0.000%", r#""ID-"000000"#];
        let started = Instant::now();
        for (idx, row) in row_targets.into_iter().enumerate() {
            editor.set_number_format(&format!("{row}:"), row_formats[idx % row_formats.len()])?;
        }
        details_for_run.borrow_mut().push(BenchDetail {
            label: "full_rows",
            elapsed_ms: as_millis(started.elapsed()),
        });

        let col_formats = [
            "#,##0.00",
            "yyyy-mm-dd",
            "@",
            r#"[Blue]#,##0.000;[Red]-#,##0.000;0.000;@"#,
        ];
        let started = Instant::now();
        for col in 0..cols.min(8) {
            let name = excel_col_name(col);
            editor.set_number_format(&format!("{name}:"), col_formats[col as usize % col_formats.len()])?;
        }
        details_for_run.borrow_mut().push(BenchDetail {
            label: "full_cols",
            elapsed_ms: as_millis(started.elapsed()),
        });
        Ok(())
    })?;
    Ok(BenchResult {
        details: Rc::try_unwrap(details).unwrap().into_inner(),
        ..result
    })
}

fn bench_switch_storm(
    test_dir: &Path,
    out_dir: &Path,
    cfg: &ScenarioConfig,
) -> Result<BenchResult> {
    println!("run switch_storm");
    let src = staged_copy(test_dir, out_dir, "test_new_ws.xlsx", "switch_storm_input.xlsx")?;
    let dst = out_dir.join("switch_storm_out.xlsx");
    run_bench("switch_storm", &src, &dst, move |editor| {
        editor.add_worksheet("BenchA")?;
        editor.add_worksheet("BenchB")?;
        editor.add_worksheet("BenchC")?;
        editor.add_worksheet("BenchD")?;

        for idx in 0..cfg.switches {
            let target = match idx % 5 {
                0 => "Sheet1",
                1 => "BenchA",
                2 => "BenchB",
                3 => "BenchC",
                _ => "BenchD",
            };
            editor.with_worksheet(target)?;
            let coord = format!("A{}", idx + 1);
            editor.set_cell(&coord, format!("switch_{idx}"))?;
        }
        Ok(())
    })
}

fn prepare_large_sheet_input(
    test_dir: &Path,
    dst: &Path,
    rows: usize,
    cols: usize,
    label: &str,
    salt: u64,
) -> Result<()> {
    let src = test_dir.join("test_new_ws.xlsx");
    fs::copy(&src, dst)
        .with_context(|| format!("failed to copy {} to {}", src.display(), dst.display()))?;
    let sheets = scan(dst)?;
    let sheet = sheets
        .first()
        .cloned()
        .context("input workbook contains no worksheets")?;
    let mut editor = XlsxEditor::open(dst, &sheet)?;
    println!("{label}: build {rows} rows x {cols} cols");
    editor.append_table(GeneratedRows::new(rows, cols, salt))?;
    save_overwrite(&mut editor, dst)?;
    Ok(())
}

fn prepare_style_input(test_dir: &Path, dst: &Path, cfg: &ScenarioConfig) -> Result<()> {
    prepare_large_sheet_input(
        test_dir,
        dst,
        cfg.style_seed_rows,
        cfg.style_cols,
        "style input",
        303,
    )?;

    let sheets = scan(dst)?;
    let sheet = sheets
        .first()
        .cloned()
        .context("style input workbook contains no worksheets")?;
    let mut editor = XlsxEditor::open(dst, &sheet)?;
    editor.add_worksheet("StyleAux")?;
    editor.append_table(GeneratedRows::new(
        (cfg.style_seed_rows / 4).max(1),
        cfg.style_cols,
        404,
    ))?;
    save_overwrite(&mut editor, dst)?;
    Ok(())
}

fn run_bench<F>(scenario: &'static str, src: &Path, dst: &Path, mut mutate: F) -> Result<BenchResult>
where
    F: FnMut(&mut XlsxEditor) -> Result<()>,
{
    let sheets = scan(src)?;
    let sheet = sheets
        .first()
        .cloned()
        .context("workbook contains no worksheets")?;

    let open_started = Instant::now();
    let mut editor = XlsxEditor::open(src, &sheet)?;
    let open_elapsed = open_started.elapsed();

    let mutate_started = Instant::now();
    mutate(&mut editor)?;
    let mutate_elapsed = mutate_started.elapsed();

    let save_started = Instant::now();
    editor.save(dst)?;
    let save_elapsed = save_started.elapsed();

    let total = open_elapsed + mutate_elapsed + save_elapsed;
    let output_bytes = fs::metadata(dst)
        .with_context(|| format!("failed to read metadata for {}", dst.display()))?
        .len();

    Ok(BenchResult {
        scenario,
        open_ms: as_millis(open_elapsed),
        mutate_ms: as_millis(mutate_elapsed),
        save_ms: as_millis(save_elapsed),
        total_ms: as_millis(total),
        output_bytes,
        input_path: src.to_path_buf(),
        output_path: dst.to_path_buf(),
        details: Vec::new(),
    })
}

fn print_report(results: &[BenchResult], cfg: &ScenarioConfig) {
    println!("# excelsior stress benchmark");
    println!();
    println!(
        "preset={} | append_rows={} append_cols={} overlap_seed_rows={} overlap_patch_rows={} style_rows={} style_cols={} style_waves={} switches={}",
        cfg.preset_name,
        cfg.append_rows,
        cfg.append_cols,
        cfg.overlap_seed_rows,
        cfg.overlap_patch_rows,
        cfg.style_seed_rows,
        cfg.style_cols,
        cfg.style_waves,
        cfg.switches
    );
    println!();
    println!("| scenario | open_ms | mutate_ms | save_ms | total_ms | output_bytes | output |");
    println!("| --- | ---: | ---: | ---: | ---: | ---: | --- |");
    for item in results {
        println!(
            "| {} | {} | {} | {} | {} | {} | {} |",
            item.scenario,
            format_ms(item.open_ms),
            format_ms(item.mutate_ms),
            format_ms(item.save_ms),
            format_ms(item.total_ms),
            item.output_bytes,
            item.output_path.display()
        );
        for detail in &item.details {
            println!("  - {}: {}", detail.label, format_ms(detail.elapsed_ms));
        }
    }
}

fn write_markdown_report(out_dir: &Path, results: &[BenchResult], cfg: &ScenarioConfig) -> Result<()> {
    let mut md = String::new();
    md.push_str("# excelsior stress benchmark\n\n");
    md.push_str("## Config\n\n");
    md.push_str(&format!(
        "- preset: `{}`\n- append_rows: `{}`\n- append_cols: `{}`\n- overlap_seed_rows: `{}`\n- overlap_patch_rows: `{}`\n- style_rows: `{}`\n- style_cols: `{}`\n- style_waves: `{}`\n- switches: `{}`\n\n",
        cfg.preset_name,
        cfg.append_rows,
        cfg.append_cols,
        cfg.overlap_seed_rows,
        cfg.overlap_patch_rows,
        cfg.style_seed_rows,
        cfg.style_cols,
        cfg.style_waves,
        cfg.switches
    ));
    md.push_str("## Results\n\n");
    md.push_str("| scenario | open_ms | mutate_ms | save_ms | total_ms | output_bytes |\n");
    md.push_str("| --- | ---: | ---: | ---: | ---: | ---: |\n");
    for item in results {
        md.push_str(&format!(
            "| {} | {} | {} | {} | {} | {} |\n",
            item.scenario,
            format_ms(item.open_ms),
            format_ms(item.mutate_ms),
            format_ms(item.save_ms),
            format_ms(item.total_ms),
            item.output_bytes
        ));
    }

    md.push_str("\n## Stage Details\n\n");
    for item in results {
        if item.details.is_empty() {
            continue;
        }
        md.push_str(&format!("### {}\n\n", item.scenario));
        for detail in &item.details {
            md.push_str(&format!("- {}: `{}` ms\n", detail.label, format_ms(detail.elapsed_ms)));
        }
        md.push('\n');
    }

    md.push_str("\n## Files\n\n");
    for item in results {
        md.push_str(&format!(
            "- `{}` input: `{}`\n- `{}` output: `{}`\n",
            item.scenario,
            item.input_path.display(),
            item.scenario,
            item.output_path.display()
        ));
    }

    let report_path = out_dir.join("benchmark-report.md");
    fs::write(&report_path, md)
        .with_context(|| format!("failed to write {}", report_path.display()))?;
    Ok(())
}

fn staged_copy(test_dir: &Path, out_dir: &Path, src_name: &str, dst_name: &str) -> Result<PathBuf> {
    let src = test_dir.join(src_name);
    let dst = out_dir.join(dst_name);
    fs::copy(&src, &dst)
        .with_context(|| format!("failed to copy {} to {}", src.display(), dst.display()))?;
    Ok(dst)
}

fn save_overwrite(editor: &mut XlsxEditor, dst: &Path) -> Result<()> {
    let tmp = dst.with_extension("tmp.xlsx");
    if tmp.exists() {
        fs::remove_file(&tmp)
            .with_context(|| format!("failed to remove stale temp file {}", tmp.display()))?;
    }
    editor.save(&tmp)?;
    if dst.exists() {
        fs::remove_file(dst)
            .with_context(|| format!("failed to remove old file {}", dst.display()))?;
    }
    fs::rename(&tmp, dst)
        .with_context(|| format!("failed to rename {} to {}", tmp.display(), dst.display()))?;
    Ok(())
}

fn excel_col_name(mut idx: u32) -> String {
    let mut out = String::new();
    loop {
        let rem = idx % 26;
        out.insert(0, (b'A' + rem as u8) as char);
        if idx < 26 {
            break;
        }
        idx = idx / 26 - 1;
    }
    out
}

fn workspace_root() -> Result<PathBuf> {
    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    manifest_dir
        .parent()
        .map(Path::to_path_buf)
        .context("failed to resolve workspace root")
}

fn unix_timestamp() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0)
}

fn as_millis(duration: Duration) -> f64 {
    duration.as_secs_f64() * 1000.0
}

fn format_ms(value: f64) -> String {
    format!("{value:.3}")
}
