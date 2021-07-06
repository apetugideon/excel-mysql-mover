use mysql::*;
use mysql::prelude::*;
use office::{Excel, DataType};
use std::env;

//#[derive(Debug, PartialEq, PartialOrd)]
fn main() {
    let args: Vec<String> = env::args().collect();
    let excel_path = &args[1];
    let db_path = &args[2];
    let tablename = &args[3];

    let mut fixed_column_fields = String::from("");
    let mut fixed_column_values = String::from("");
    if args.len() > 4 {
        let fixed_values = &args[4];
        let fixed_columns = fixed_values.split(";");
        for fixed_column in fixed_columns {
            let fixed_col_val = fixed_column.split(":");
            let fixed_col_val: Vec<&str> = fixed_col_val.collect();
            let column_name = fixed_col_val[0];
            let column_value = fixed_col_val[1];
    
            let formated_field = format!(", {}", column_name);
            fixed_column_fields.push_str(&formated_field);
    
            let formated_values = format!(", {:?}", column_value);
            fixed_column_values.push_str(&formated_values);
        }
    }


    let pool = Pool::new(db_path).unwrap();
    let mut conn = pool.get_conn().unwrap();

    
    let mut ccount = 0;
    let mut rcount: i64 = 1;
    let file_name = String::from(excel_path);
    let mut excel = Excel::open(file_name).unwrap();
    let r = excel.worksheet_range("Sheet1").unwrap();


    let mut column_names = Vec::new();
    let mut field_string = String::from("");


    for row in r.rows() {
        if ccount == 0 { ccount = row.len(); }
        if rcount == 1 {
            for index in row {
                let current_cols = String::from(get_string_value(index));
                let mut column_string = format!("{}", current_cols);
                if field_string != "" {
                    column_string = format!(", {}", current_cols);
                }
                field_string.push_str(&column_string);
                column_names.push(current_cols);
            }

            if fixed_column_fields != "" {
                field_string.push_str(&fixed_column_fields);
            }
        } else {
            let mut colcnt = 0;
            let mut value_string = String::from("");
            for _col_name in &column_names {
                let column_value = extract_value(&row[colcnt]);
                let mut value_str = format!("{:?}", column_value);
                if value_string != "" {
                    value_str = format!(", {:?}", column_value);
                }
                value_string.push_str(&value_str);
                colcnt += 1;
            }

            if fixed_column_values != "" {
                value_string.push_str(&fixed_column_values);
            }
            
            let q_string = format!("insert into {} ({}) values ({})", tablename, field_string, value_string);
            let stmt = conn.prep(q_string).unwrap();
            conn.exec_drop(&stmt, ()).unwrap();
        }
        rcount = rcount + 1;
    }
}


fn extract_value(inval: &DataType) -> String {
    match inval {
        DataType::String(str) => String::from(str),
        DataType::Int(inum) => inum.to_string(),
        DataType::Float(fnum) => fnum.to_string(),
        DataType::Bool(bool_val) => bool_val.to_string(),
        _ => "".to_string()
    }
}


fn get_string_value(inval: &DataType) -> &str {
    let matched_value = match inval {
        DataType::String(str) => str,
        _ => "",
    };
    return matched_value;
}