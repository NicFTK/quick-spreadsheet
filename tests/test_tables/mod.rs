use umya_spreadsheet::{self, new_file, writer, Comment, Coordinate};

/// Excel表格位置信息
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlTable {
    /// 唯一标识符，不带<TableID>标签
    pub uuid: String,

    /// 开始位置地址
    pub start_cell: Coordinate,
    /// 结束位置地址
    pub end_cell: Coordinate,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct CellIndex {
    pub row: usize,
    pub col: usize,
}

#[derive(Debug, Clone)]
pub struct Row {
    pub cells: Vec<String>, // For simplicity, storing cell values as String
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct MergedCells {
    /// 合并的起始位置
    /// 从(0, 0)开始
    /// row, col
    pub start: CellIndex,
    /// 合并的结束位置
    /// 从(0, 0)开始
    pub end: CellIndex,

    /// 跨行合并的行数
    pub row_span: usize,

    /// 跨列合并的列数
    pub col_span: usize,
}

#[derive(Debug, Clone)]
/// 表格位置矩阵
/// 用于表示表格的结构和位置关系;
/// 是一个完整的矩形矩阵；
/// 构造完成后，内部不会存在一行完全都被合并的
pub struct TableMatrix {
    /// 矩阵的行数
    pub rows: usize,
    /// 矩阵的列数
    pub cols: usize,
    /// 矩阵数据
    pub data: Vec<Row>,
    ///记录合并单元格的相关信息
    pub merged_cells: Vec<MergedCells>,
}

#[test]
fn create_test_table_file() {
    let mut book = new_file();
    let sheet = book.get_sheet_mut(&0).unwrap();
    sheet.set_name("Sheet1");

    // Table 1
    let mut comment_start = Comment::default();
    comment_start.new_comment("B2");
    comment_start.set_text_string("<TableStart><TableID>table1</TableID>");
    sheet.add_comments(comment_start);

    let mut comment_end = Comment::default();
    comment_end.new_comment("D5");
    comment_end.set_text_string("<TableEnd>");
    sheet.add_comments(comment_end);

    // Table 2
    let mut comment_start_2 = Comment::default();
    comment_start_2.new_comment("F2");
    comment_start_2.set_text_string("<TableStart><TableID>table2</TableID>");
    sheet.add_comments(comment_start_2);

    let mut comment_end_2 = Comment::default();
    comment_end_2.new_comment("H8");
    comment_end_2.set_text_string("<TableEnd>");
    sheet.add_comments(comment_end_2);

    // Add a merged cell for testing
    sheet.add_merge_cells("B3:C4");

    let path = std::path::Path::new(r"G:\Aduit_Project\【测试文件夹】\07-xlsx读取生成\table_test.xlsx");
    let _ = writer::xlsx::write(&book, path);
}

#[test]
fn find_tables_in_sheet() {
    let path = std::path::Path::new(r"G:\Aduit_Project\【测试文件夹】\07-xlsx读取生成\table_test.xlsx");
    let book = umya_spreadsheet::reader::xlsx::read(path).unwrap();
    let sheet = book.get_sheet_by_name("Sheet1").unwrap();

    let mut comments: Vec<_> = sheet.get_comments().iter().cloned().collect();
    comments.sort_by(|a, b| {
        a.get_coordinate().cmp(b.get_coordinate())
    });

    let mut tables = Vec::new();
    let mut start_stack: Vec<(String, Coordinate)> = Vec::new();

    for comment in &comments {
        if let Some(text) = comment.get_text().get_text() {
            println!("Comment at {}",  text.get_value());
            let text_val = text.get_value();
            if text_val.contains("<TableStart>") {
                if let Some(uuid) = extract_uuid(text_val) {
                    start_stack.push((uuid, comment.get_coordinate().clone()));
                }
            } else if text_val.contains("<TableEnd>") {
                if let Some((uuid, start_cell)) = start_stack.pop() {
                    tables.push(XlTable {
                        uuid,
                        start_cell,
                        end_cell: comment.get_coordinate().clone(),
                    });
                }
            }
        }else{
            println!("Comment at {:?} has no text", comment.get_coordinate());
        }
    };

    for table in &tables {
        println!("Found table: UUID: {}, Start: {:?}, End: {:?}", table.uuid, table.start_cell, table.end_cell);
    }


    // 构造内容矩阵
    let mut table_matrices = Vec::new();
    for table in &tables {
        let start_col = *table.start_cell.get_col_num();
        let start_row = *table.start_cell.get_row_num();
        let end_col = *table.end_cell.get_col_num();
        let end_row = *table.end_cell.get_row_num();

        let rows = (end_row - start_row + 1) as usize;
        let cols = (end_col - start_col + 1) as usize;

        let mut data = Vec::with_capacity(rows);
        for r in start_row..=end_row {
            let mut row_data = Vec::with_capacity(cols);
            for c in start_col..=end_col {
                let cell_value = sheet.get_value((c, r));
                row_data.push(cell_value);
            }
            data.push(Row { cells: row_data });
        }

        let merged_cells_in_table: Vec<MergedCells> = sheet
            .get_merge_cells()
            .iter()
            .filter_map(|merged_range| {
                let merge_start_col = merged_range.get_coordinate_start_col().unwrap().get_num();
                let merge_start_row = merged_range.get_coordinate_start_row().unwrap().get_num();
                let merge_end_col = merged_range.get_coordinate_end_col().unwrap().get_num();
                let merge_end_row = merged_range.get_coordinate_end_row().unwrap().get_num();

                if *merge_start_col >= start_col && *merge_end_col <= end_col &&
                   *merge_start_row >= start_row && *merge_end_row <= end_row {
                    
                    let start = CellIndex {
                        row: (*merge_start_row - start_row) as usize,
                        col: (*merge_start_col - start_col) as usize,
                    };
                    let end = CellIndex {
                        row: (*merge_end_row - start_row) as usize,
                        col: (*merge_end_col - start_col) as usize,
                    };
                    let row_span = (*merge_end_row - *merge_start_row + 1) as usize;
                    let col_span = (*merge_end_col - *merge_start_col + 1) as usize;

                    Some(MergedCells {
                        start,
                        end,
                        row_span,
                        col_span,
                    })
                } else {
                    None
                }
            })
            .collect();

        let table_matrix = TableMatrix {
            rows,
            cols,
            data,
            merged_cells: merged_cells_in_table,
        };
        table_matrices.push(table_matrix);
    }
    println!("Found {} tables", tables.len());
    // Debug output
    for (i, table) in tables.iter().enumerate() {
        println!("Table {}: UUID: {}, Start: {:?}, End: {:?}", i + 1, table.uuid, table.start_cell, table.end_cell);
        let matrix = &table_matrices[i];
        for row in &matrix.data {
            println!("{:?}", row.cells);
        }
        println!("Merged Cells: {:?}", matrix.merged_cells);
    }

}

fn extract_uuid(text: &str) -> Option<String> {
    let start_tag = "<TableID>";
    let end_tag = "</TableID>";
    text.find(start_tag)
        .and_then(|start| {
            let start = start + start_tag.len();
            text[start..].find(end_tag).map(|end| text[start..start + end].to_string())
        })
}

#[test]
fn read_and_write_test() {
    // TODO: a.xlsx
    let path = std::path::Path::new(r"G:\Aduit_Project\【测试文件夹】\03-报告更新v2\B公司报告 - 副本.docx_20250805092655.xlsx");
    let book = umya_spreadsheet::reader::xlsx::read(path).unwrap();

    // TODO: b.xlsx
    let path = std::path::Path::new(r"G:\Aduit_Project\【测试文件夹】\03-报告更新v2\B公司报告 - 副本.docx_write.xlsx");
    let _ = writer::xlsx::write(&book, path);
}