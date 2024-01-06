use calamine::{open_workbook, Reader, Xls};
use chrono::NaiveDate;
use csv::Writer;
use std::error::Error;
use std::path::Path;

pub fn main() -> Result<(), Box<dyn Error>> {
    let file_path = "../cartolas/cartola.xls";
    let path = Path::new(file_path);
    let mut reader: Xls<_> = open_workbook(path)?;
    let r = match reader.worksheet_range("Hoja1") {
        Ok(range) => range,
        Err(err) => return Err(Box::new(err)),
    };
    let mut i = 0;
    let mut records: Vec<Vec<String>> = vec![];
    let mut writer = Writer::from_path("output.csv")?;

    for row in r.rows() {
        i += 1;
        if i < 21 { continue; } // Skip headers

        // Skip empty rows following the transaction list
        if row[0].is_empty() { break; }

        let mut record: Vec<String> = vec![];

        for (j, cell) in row.iter().enumerate() {
            if j == 2 { continue; } // Skip column 2

            let cell = cell.to_string();

            if cell.contains("SALDO INICIAL") || cell.contains("SALDO FINAL") {
                continue;
            }

            if j == 0 { // Parse date
                let date = NaiveDate::parse_from_str(&cell, "%d/%m/%Y")?;
                record.push(date.to_string());
            } else {
                record.push(cell);
            }
        }

        records.push(record);
    }

    // Sort records by date
    records.sort_by(|a, b| a[0].cmp(&b[0]));

    // Write records to CSV
    for record in records {
        writer.write_record(&record)?;
    }

    writer.flush()?;

    Ok(())
}