//! Read saved cell values and formulas with a plain xlsx reader (no recalculation):
//! the web client's engine writes each formula's cached result when it saves.
use crate::{
    domain::schemas::CellValue,
    error::{AppError, Result},
};
use calamine::{Data, Range, Reader, Xlsx};
use std::io::Cursor;

/// Most cells one read returns, so an agent cannot ask for a whole 1M-row sheet.
pub const MAX_CELLS: usize = 5000;

pub struct OpenWorkbook {
    workbook: Xlsx<Cursor<Vec<u8>>>,
}

pub struct SheetRead {
    pub sheet: String,
    pub sheets: Vec<String>,
    pub range: String,
    pub cells: Vec<CellValue>,
}

fn invalid(what: &str) -> AppError {
    AppError::Invalid(what.into())
}

pub fn col_name(mut col: u32) -> String {
    let mut name = Vec::new();
    col += 1;
    while col > 0 {
        let rem = (col - 1) % 26;
        name.push(b'A' + u8::try_from(rem).unwrap_or(0));
        col = (col - 1) / 26;
    }
    name.reverse();
    String::from_utf8(name).unwrap_or_default()
}

pub fn address(row: u32, col: u32) -> String {
    format!("{}{}", col_name(col), row + 1)
}

/// `B7` → (row 6, col 1); `$B$7` is accepted.
pub fn parse_address(text: &str) -> Option<(u32, u32)> {
    let text = text.trim().replace('$', "").to_ascii_uppercase();
    let split = text.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = text.split_at(split);
    if letters.is_empty() || letters.len() > 3 || !letters.chars().all(|c| c.is_ascii_uppercase()) {
        return None;
    }
    let col = letters.bytes().fold(0u32, |acc, b| acc * 26 + u32::from(b - b'A' + 1)) - 1;
    let row: u32 = digits.parse().ok()?;
    if row == 0 || row > 1_048_576 || col >= 16_384 {
        return None;
    }
    Some((row - 1, col))
}

fn parse_range(text: &str) -> Option<((u32, u32), (u32, u32))> {
    let mut parts = text.split(':');
    let start = parse_address(parts.next()?)?;
    let end = parts.next().map_or(Some(start), parse_address)?;
    if parts.next().is_some() {
        return None;
    }
    Some(((start.0.min(end.0), start.1.min(end.1)), (start.0.max(end.0), start.1.max(end.1))))
}

/// Text the grader and agents compare: integers without `.0`, booleans as TRUE/FALSE.
pub fn display(data: &Data) -> (String, Option<f64>) {
    match data {
        Data::Empty => (String::new(), None),
        Data::String(text) => (text.clone(), None),
        Data::Float(number) => (format_number(*number), Some(*number)),
        #[allow(clippy::cast_precision_loss)]
        Data::Int(number) => (number.to_string(), Some(*number as f64)),
        Data::Bool(flag) => ((if *flag { "TRUE" } else { "FALSE" }).into(), None),
        Data::Error(error) => (format!("{error}"), None),
        Data::DateTime(date) => (date.as_f64().to_string(), Some(date.as_f64())),
        Data::DateTimeIso(text) | Data::DurationIso(text) => (text.clone(), None),
    }
}

fn format_number(number: f64) -> String {
    if number.fract() == 0.0 && number.abs() < 1e15 {
        format!("{number:.0}")
    } else {
        number.to_string()
    }
}

impl OpenWorkbook {
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Xlsx::new(Cursor::new(bytes))
            .map(|workbook| Self { workbook })
            .map_err(|_| invalid("workbook"))
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.workbook.sheet_names()
    }

    fn resolve_sheet(&self, sheet: Option<&str>) -> Result<String> {
        let names = self.sheet_names();
        match sheet {
            None => names.into_iter().next().ok_or_else(|| invalid("sheet")),
            Some(wanted) => names
                .into_iter()
                .find(|name| name.eq_ignore_ascii_case(wanted.trim()))
                .ok_or_else(|| AppError::NotFound(format!("sheet {wanted}"))),
        }
    }

    fn values(&mut self, sheet: &str) -> Result<Range<Data>> {
        self.workbook.worksheet_range(sheet).map_err(|_| invalid("sheet"))
    }

    /// One cell's cached value; empty when the cell is blank.
    pub fn cell(&mut self, sheet: Option<&str>, cell: &str) -> Result<(String, Option<f64>)> {
        let sheet = self.resolve_sheet(sheet)?;
        let (row, col) = parse_address(cell).ok_or_else(|| invalid("cell"))?;
        let values = self.values(&sheet)?;
        Ok(values.get_value((row, col)).map_or((String::new(), None), display))
    }

    pub fn read(&mut self, sheet: Option<&str>, range: Option<&str>) -> Result<SheetRead> {
        let sheets = self.sheet_names();
        let sheet = self.resolve_sheet(sheet)?;
        let values = self.values(&sheet)?;
        let formulas = self.workbook.worksheet_formula(&sheet).unwrap_or_default();
        let bounds = match range {
            Some(text) => parse_range(text).ok_or_else(|| invalid("range"))?,
            None => match (values.start(), values.end()) {
                (Some(start), Some(end)) => (start, end),
                _ => ((0, 0), (0, 0)),
            },
        };
        let ((r1, c1), (r2, c2)) = bounds;
        let mut cells = Vec::new();
        'rows: for row in r1..=r2 {
            for col in c1..=c2 {
                let value = values.get_value((row, col)).map_or((String::new(), None), display);
                let formula = formulas
                    .get_value((row, col))
                    .filter(|formula| !formula.is_empty())
                    .map(|formula| format!("={formula}"));
                if value.0.is_empty() && formula.is_none() {
                    continue;
                }
                cells.push(CellValue {
                    address: address(row, col),
                    value: value.0,
                    number: value.1,
                    formula,
                });
                if cells.len() >= MAX_CELLS {
                    break 'rows;
                }
            }
        }
        let range = format!("{}:{}", address(r1, c1), address(r2, c2));
        Ok(SheetRead { sheet, sheets, range, cells })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn addresses_round_trip() {
        for (text, rc) in [
            ("A1", (0, 0)),
            ("Z9", (8, 25)),
            ("AA10", (9, 26)),
            ("$XFD$1048576", (1_048_575, 16_383)),
        ] {
            assert_eq!(parse_address(text), Some(rc));
            assert_eq!(address(rc.0, rc.1), text.replace('$', ""));
        }
        for bad in ["", "1A", "A0", "XFE1", "AAAA1", "A1048577"] {
            assert_eq!(parse_address(bad), None, "{bad}");
        }
        assert_eq!(parse_range("C3:A1"), Some(((0, 0), (2, 2))));
    }
}
