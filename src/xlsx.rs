//! Tiny OOXML (.xlsx) writer — just enough to emit a multi-sheet workbook with
//! inline strings, numbers, a header style and a couple of number formats.
//! Inline strings keep us free of a shared-string table; styles are a fixed
//! tiny stylesheet referenced by index.

use std::io::Write;
use zip::write::FileOptions;

/// A single cell value.
#[derive(Clone)]
pub enum Cell {
    Empty,
    Text(String),
    Num(f64),
}

pub fn t<S: Into<String>>(s: S) -> Cell {
    Cell::Text(s.into())
}
pub fn n(v: f64) -> Cell {
    Cell::Num(v)
}
pub fn e() -> Cell {
    Cell::Empty
}

/// Style indices into the fixed stylesheet emitted by `styles_xml`.
const S_DEFAULT: u32 = 0;
const S_HEADER: u32 = 1;
const S_TITLE: u32 = 2;
const S_MONEY: u32 = 3;

pub struct Sheet {
    pub name: String,
    pub rows: Vec<Vec<Cell>>,
    /// row index (0-based) -> force header style across the row.
    pub header_rows: Vec<usize>,
    pub title_rows: Vec<usize>,
    pub col_widths: Vec<f64>,
}

impl Sheet {
    pub fn new(name: &str) -> Self {
        Sheet {
            name: name.to_string(),
            rows: Vec::new(),
            header_rows: Vec::new(),
            title_rows: Vec::new(),
            col_widths: Vec::new(),
        }
    }
    pub fn row(&mut self, cells: Vec<Cell>) -> &mut Self {
        self.rows.push(cells);
        self
    }
    pub fn header(&mut self, cells: Vec<Cell>) -> &mut Self {
        self.header_rows.push(self.rows.len());
        self.rows.push(cells);
        self
    }
    pub fn title(&mut self, text: &str) -> &mut Self {
        self.title_rows.push(self.rows.len());
        self.rows.push(vec![t(text)]);
        self
    }
    pub fn widths(&mut self, w: &[f64]) -> &mut Self {
        self.col_widths = w.to_vec();
        self
    }
}

fn col_letter(mut c: usize) -> String {
    let mut s = String::new();
    c += 1;
    while c > 0 {
        let r = (c - 1) % 26;
        s.insert(0, (b'A' + r as u8) as char);
        c = (c - 1) / 26;
    }
    s
}

fn esc(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn sheet_xml(sheet: &Sheet) -> String {
    let mut x = String::new();
    x.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    x.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#);
    if !sheet.col_widths.is_empty() {
        x.push_str("<cols>");
        for (i, w) in sheet.col_widths.iter().enumerate() {
            x.push_str(&format!(
                r#"<col min="{0}" max="{0}" width="{1}" customWidth="1"/>"#,
                i + 1,
                w
            ));
        }
        x.push_str("</cols>");
    }
    x.push_str("<sheetData>");
    for (ri, row) in sheet.rows.iter().enumerate() {
        let rnum = ri + 1;
        let is_header = sheet.header_rows.contains(&ri);
        let is_title = sheet.title_rows.contains(&ri);
        x.push_str(&format!(r#"<row r="{}">"#, rnum));
        for (ci, cell) in row.iter().enumerate() {
            let cref = format!("{}{}", col_letter(ci), rnum);
            let style = if is_title {
                S_TITLE
            } else if is_header {
                S_HEADER
            } else {
                match cell {
                    Cell::Num(_) => S_MONEY,
                    _ => S_DEFAULT,
                }
            };
            match cell {
                Cell::Empty => {
                    x.push_str(&format!(r#"<c r="{}" s="{}"/>"#, cref, style));
                }
                Cell::Text(s) => {
                    x.push_str(&format!(
                        r#"<c r="{}" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>"#,
                        cref,
                        style,
                        esc(s)
                    ));
                }
                Cell::Num(v) => {
                    x.push_str(&format!(
                        r#"<c r="{}" s="{}"><v>{}</v></c>"#,
                        cref, style, v
                    ));
                }
            }
        }
        x.push_str("</row>");
    }
    x.push_str("</sheetData></worksheet>");
    x
}

fn styles_xml() -> String {
    // fonts: 0 normal, 1 bold, 2 bold white (title)
    // fills: 0 none, 1 gray125, 2 header blue, 3 title dark
    // numFmt 164: #,##0.00######
    r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00######;[Red]-#,##0.00######"/></numFmts>
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0B3D2E"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"##.to_string()
}

/// Write all sheets to an .xlsx at `path`.
pub fn write_xlsx(path: &str, sheets: &[Sheet]) -> std::io::Result<()> {
    let file = std::fs::File::create(path)?;
    let mut zip = zip::ZipWriter::new(file);
    let opts: FileOptions =
        FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", opts)?;
    let mut ct = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    );
    for i in 0..sheets.len() {
        ct.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
            i + 1
        ));
    }
    ct.push_str("</Types>");
    zip.write_all(ct.as_bytes())?;

    // _rels/.rels
    zip.start_file("_rels/.rels", opts)?;
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#)?;

    // xl/workbook.xml
    zip.start_file("xl/workbook.xml", opts)?;
    let mut wb = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"#,
    );
    for (i, s) in sheets.iter().enumerate() {
        wb.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            esc(&s.name),
            i + 1,
            i + 1
        ));
    }
    wb.push_str("</sheets></workbook>");
    zip.write_all(wb.as_bytes())?;

    // xl/_rels/workbook.xml.rels
    zip.start_file("xl/_rels/workbook.xml.rels", opts)?;
    let mut rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for i in 0..sheets.len() {
        rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
            i + 1,
            i + 1
        ));
    }
    let styles_rid = sheets.len() + 1;
    rels.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
        styles_rid
    ));
    rels.push_str("</Relationships>");
    zip.write_all(rels.as_bytes())?;

    // styles
    zip.start_file("xl/styles.xml", opts)?;
    zip.write_all(styles_xml().as_bytes())?;

    // sheets
    for (i, s) in sheets.iter().enumerate() {
        zip.start_file(format!("xl/worksheets/sheet{}.xml", i + 1), opts)?;
        zip.write_all(sheet_xml(s).as_bytes())?;
    }

    zip.finish()?;
    Ok(())
}
