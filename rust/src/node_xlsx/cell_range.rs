#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRange {
    pub start_row: u32,
    pub end_row: u32,
    pub start_col: u32,
    pub end_col: u32,
}

impl CellRange {
    pub fn new(start_row: u32, end_row: u32, start_col: u32, end_col: u32) -> Result<Self, String> {
        if start_row > end_row {
            return Err("start_row must be less than or equal to end_row".to_string());
        }
        if start_col > end_col {
            return Err("start_col must be less than or equal to end_col".to_string());
        }
        if end_row > 1048576 {
            return Err("end_row must be less than or equal to 1048576".to_string());
        }
        if end_col > 16384 {
            return Err("end_col must be less than or equal to 16384".to_string());
        }
        if start_row == end_row && start_col == end_col {
            return Err(
                "start_row and end_row cannot be the same if start_col and end_col are the same"
                    .to_string(),
            );
        }

        return Ok(Self {
            start_row,
            end_row,
            start_col,
            end_col,
        });
    }

    pub fn overlaps(&self, other: &Self) -> bool {
        self.start_row <= other.end_row
            && self.end_row >= other.start_row
            && self.start_col <= other.end_col
            && self.end_col >= other.start_col
    }
}
