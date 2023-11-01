#![windows_subsystem = "windows"]
use std::fs;
use docx_rs::*;
use std::path::PathBuf;
use native_dialog::FileDialog;

pub fn create_table(path: &std::fs::DirEntry, dir_path: &str) -> Table {
    let table = Table::new(vec![
        TableRow::new(vec![
            TableCell::new()
                .grid_span(2)
                .shading(Shading::new().color("006666").shd_type(ShdType::Solid))
                .add_paragraph(Paragraph::new()
                    .line_spacing(LineSpacing::new().before(0).after(0))
                    .add_hyperlink(Hyperlink::new(path.path().into_os_string().into_string().unwrap(), HyperlinkType::External)
                        .add_run(Run::new().add_text(path.path().file_name().unwrap().to_str().unwrap()).bold().color("FFFFFF").fonts(RunFonts::new().ascii("Calibri")))
                    )
                )
            ]),
        ])
        .align(TableAlignmentType::Center)
        .set_grid(vec![4249, 4249])
        .margins(TableCellMargins::new().margin(25, 25, 25, 25))
        .layout(TableLayoutType::Fixed);

    add_row_to_table(&table, path, dir_path)
}

pub fn add_row_to_table(table: &Table, entry: &std::fs::DirEntry, dir_path: &str) -> Table {
    let paths = fs::read_dir(entry.path().into_os_string().into_string().unwrap()).unwrap();

    let files: Vec<PathBuf> = paths
        .filter_map(|path| path.ok().map(|p| p.path()))
        .filter(|p| p.is_file())
        .collect();

    let mut t = table.clone();
   
    if files.len() > 0 {
        let mut table_cell = TableCell::new();
        for file in files {
            let f_os_string = file.clone().into_os_string().into_string().unwrap();
            let f_file_name = file.file_name().unwrap().to_str().unwrap();
            table_cell = table_cell.add_paragraph(Paragraph::new()
                .line_spacing(LineSpacing::new().before(0).after(0))
                .add_hyperlink(Hyperlink::new(f_os_string.clone(), HyperlinkType::External)
                    .add_run(Run::new().add_text(f_file_name).fonts(RunFonts::new().ascii("Calibri")))
                )
            );
        }

        let path = entry.path();
        let relative_path = path.strip_prefix(dir_path.to_string()).unwrap_or(path.as_path());
        if relative_path.to_str().unwrap().contains(std::path::MAIN_SEPARATOR) {
            let (last, rest) = relative_path.to_str().unwrap().rsplit_once(std::path::MAIN_SEPARATOR).unwrap_or(("", ""));

            t = t.add_row(
                TableRow::new(vec![ 
                   TableCell::new()
                    .add_paragraph(Paragraph::new()
                    .line_spacing(LineSpacing::new().before(0).after(0))
                    .add_run(Run::new().add_text(last).fonts(RunFonts::new().ascii("Calibri")))
                    .add_run(Run::new().add_text(std::path::MAIN_SEPARATOR).fonts(RunFonts::new().ascii("Calibri")))
                    .add_run(Run::new().add_text(rest).bold().fonts(RunFonts::new().ascii("Calibri")))
                    ),
                    table_cell
                ])
            );
        } else {
            t = t.add_row(
                TableRow::new(vec![
                    TableCell::new()
                    .add_paragraph(Paragraph::new()
                    .line_spacing(LineSpacing::new().before(0).after(0))
                    .add_run(Run::new().add_text(relative_path.to_str().unwrap()).bold().fonts(RunFonts::new().ascii("Calibri")))
                    ),
                    table_cell
                ])
            );

        }
    }

    let paths = fs::read_dir(entry.path().into_os_string().into_string().unwrap()).unwrap();
    for path in paths {
        let p = path.as_ref().unwrap();
        let is_file = p.path().is_file();
        if !is_file {
            t = add_row_to_table(&t, &p, &dir_path);
        }
    }

    t
}

fn main() -> Result<(), DocxError> {
    let dialog = FileDialog::new()
        .set_location(r"C:\Users\Tanguy\Documents\SQL")
        .show_open_single_dir()
        .unwrap();

    let dir_path = match dialog {
        Some(path) => path,
        None => return Ok(()),
    };
    let paths = fs::read_dir(&dir_path.to_str().unwrap()).unwrap();
    let mut docx = Docx::new();

    for path in paths {
        let table = create_table(&path.unwrap(), &dir_path.to_str().unwrap());
        docx = docx.add_table(table).add_paragraph(Paragraph::new().line_spacing(LineSpacing::new().before(0).after(0)));
    }
    let path_buf: PathBuf = [&dir_path.to_str().unwrap(), "appendices.docx"].iter().collect();
    let file = std::fs::File::create(&path_buf).unwrap();
    docx.build().pack(file)?;
    Ok(())
}
