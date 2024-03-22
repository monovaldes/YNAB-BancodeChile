use calamine::{open_workbook, Reader, Xls};
use chrono::NaiveDate;
use csv::Writer;
use std::error::Error;

pub fn main() -> Result<(), Box<dyn Error>> {
    let mut records: Vec<Vec<String>> = vec![];
    let mut writer = Writer::from_path("output.csv")?;

    // Create Path vector for all xls files in cartolas directory
    let paths = std::fs::read_dir("../cartolas")?
        .map(|res| res.map(|e| e.path()))
        .collect::<Result<Vec<_>, std::io::Error>>()?;
    
    // Iterate over each file in the cartolas directory
    for path in paths {
        // Skip if file name does not end in .xls
        if path.file_name().unwrap().to_str().unwrap().chars().rev().take(4).collect::<String>() != "slx." {
            continue;
        }

        println!("Processing file: {:?}", path);

        let path = path.as_path();
        let mut reader: Xls<_> = open_workbook(path)?;
        let r = match reader.worksheet_range("Hoja1") {
            Ok(range) => range,
            Err(err) => return Err(Box::new(err)),
        };
        let mut i = 0;

        for row in r.rows() {
            i += 1;
            if i < 20 { continue; } // Skip headers
    
            // Skip empty rows following the transaction list
            if row[0].is_empty() {
                break;
            }
    
            let mut record: Vec<String> = vec![];
    
            for (j, cell) in row.iter().enumerate() {
                if j == 2 || j == 5 { continue; } // Skip columns 2 and 5
                let cell = cell.to_string();
                
                if cell.contains("SALDO INICIAL") || cell.contains("SALDO FINAL") {
                    continue;
                }
    
                if j == 0 { // Parse date
                    // get year from file name last 4 digits (excluding .xls extension), eg 1997.xls should return 1997
                    let year = path.file_name().unwrap().to_str().unwrap().chars().rev().take(8).collect::<String>().chars().rev().collect::<String>();
                    // remove .xls
                    let year = year.chars().take(4).collect::<String>();
                    // add year to cell
                    let cell = format!("{}/{}", cell, year);
                    let date = NaiveDate::parse_from_str(&cell, "%d/%m/%Y")?;
                    record.push(date.to_string());
                } else {
                    record.push(cell);
                }
            }
            // add empty string on column 2
            record.insert(2, "".to_string());

            // skip push if second column is empty
            if record[1].is_empty() {
                continue;
            }
    
            records.push(record);
        }
    }
    
    // Sort records by date
    records.sort_by(|a, b| a[0].cmp(&b[0]));

    // return if no records
    if records.len() == 0 { return Ok(()); }

    // write headers
    writer.write_record(&["Date","Payee","Memo","Outflow","Inflow"])?;

    // Write records to CSV
    for record in records {
        writer.write_record(&record)?;
    }

    // Flush writer
    println!("Flushing writer");

    writer.flush()?;

    Ok(())
}