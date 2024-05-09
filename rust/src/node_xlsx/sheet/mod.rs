mod array_formula_value;
mod conditional_format_value;

use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsObject, JsString, JsValue},
};
use rust_xlsxwriter::{Format, Worksheet, XlsxError};

use self::{
    array_formula_value::ArrayFormulaSheetValue,
    conditional_format_value::ConditionalFormatSheetValue,
};

use super::{
    cell::NodeXlsxCell, conditional_format::c_type::NodeXlsxConditionalFormatType,
    types::NodeXlsxTypes,
};

pub struct NodeXlsxSheet {
    name: String,
    cells: Vec<NodeXlsxCell>,
    conditional_formats: Vec<ConditionalFormatSheetValue>,
    array_formulas: Vec<ArrayFormulaSheetValue>,

    format_map: HashMap<u32, Format>,
    conditional_format_map: HashMap<u32, NodeXlsxConditionalFormatType>,
}

impl NodeXlsxSheet {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let name: Handle<JsString> = obj.get(cx, "name")?;
        let name = name.value(cx);

        let cells: Handle<JsArray> = obj.get(cx, "cells")?;
        let cells: Vec<Handle<JsValue>> = cells.to_vec(cx)?;

        let conditional_formats: Handle<JsArray> = obj.get(cx, "conditionalFormats")?;
        let conditional_formats: Vec<Handle<JsValue>> = conditional_formats.to_vec(cx)?;

        let array_formulas: Handle<JsArray> = obj.get(cx, "arrayFormulas")?;
        let array_formulas: Vec<Handle<JsValue>> = array_formulas.to_vec(cx)?;

        let mut inner_cells = vec![];
        let mut inner_formulas = vec![];
        let mut inner_conditional_formats = vec![];
        let mut format_map = HashMap::new();
        let mut conditional_format_map = HashMap::new();

        for formula in array_formulas {
            let formula = formula.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let formula = ArrayFormulaSheetValue::from_js_object(cx, formula, &mut format_map)?;
            inner_formulas.push(formula);
        }

        for cell in cells {
            let cell = cell.downcast_or_throw::<JsObject, FunctionContext>(cx)?;

            let cell = NodeXlsxCell::from_js_object(cx, cell, &mut format_map)?;
            inner_cells.push(cell);
        }

        for conditional_format in conditional_formats {
            let conditional_format =
                conditional_format.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let conditional_format = ConditionalFormatSheetValue::from_js_object(
                cx,
                conditional_format,
                &mut format_map,
                &mut conditional_format_map,
            )?;
            inner_conditional_formats.push(conditional_format);
        }

        return Ok(Self {
            name,
            cells: inner_cells,
            conditional_formats: inner_conditional_formats,
            array_formulas: inner_formulas,
            format_map,
            conditional_format_map,
        });
    }
}

impl TryInto<Worksheet> for NodeXlsxSheet {
    type Error = XlsxError;
    fn try_into(self) -> Result<Worksheet, XlsxError> {
        let format_map = self.format_map;
        let conditional_format_map = self.conditional_format_map;
        let mut worksheet = Worksheet::new();
        worksheet.set_name(&self.name)?;
        for cf in self.conditional_formats {
            cf.set_conditional_format(&mut worksheet, &conditional_format_map)?;
        }
        for af in self.array_formulas {
            let first_row = af.first_row;
            let last_row = af.last_row;
            let first_column = af.first_column;
            let last_column = af.last_column;
            let formula = af.formula;
            let dynamic = af.dynamic;
            let has_format = af.format.is_some();
            if dynamic && !has_format {
                worksheet.write_dynamic_array_formula(
                    first_row,
                    first_column,
                    last_row,
                    last_column,
                    formula,
                )?;
            } else if dynamic && has_format {
                let format = af.format.unwrap();
                let format = format_map.get(&format).unwrap();
                worksheet.write_dynamic_array_formula_with_format(
                    first_row,
                    first_column,
                    last_row,
                    last_column,
                    formula,
                    format,
                )?;
            } else if !dynamic && !has_format {
                worksheet.write_array_formula(
                    first_row,
                    first_column,
                    last_row,
                    last_column,
                    formula,
                )?;
            } else {
                let format = af.format.unwrap();
                let format = format_map.get(&format).unwrap();
                worksheet.write_array_formula_with_format(
                    first_row,
                    first_column,
                    last_row,
                    last_column,
                    formula,
                    format,
                )?;
            }
        }
        for cell in self.cells {
            match cell.cell_type {
                NodeXlsxTypes::String(value) | NodeXlsxTypes::Unknown(value) => {
                    if let Some(format) = cell.format {
                        let format = format_map.get(&format).unwrap();
                        worksheet.write_string_with_format(cell.row, cell.col, &value, format)?;
                    } else {
                        worksheet.write_string(cell.row, cell.col, &value)?;
                    }
                }
                NodeXlsxTypes::Number(value) => {
                    if let Some(format) = cell.format {
                        let format = format_map.get(&format).unwrap();
                        worksheet.write_number_with_format(cell.row, cell.col, value, format)?;
                    } else {
                        worksheet.write_number(cell.row, cell.col, value)?;
                    }
                }
                NodeXlsxTypes::Link(value) => {
                    if let Some(format) = cell.format {
                        let format = format_map.get(&format).unwrap();
                        worksheet.write_url_with_format(cell.row, cell.col, value, format)?;
                    } else {
                        worksheet.write_url(cell.row, cell.col, value)?;
                    }
                }
                NodeXlsxTypes::Date(value) => {
                    if let Some(format) = cell.format {
                        let format = format_map.get(&format).unwrap();
                        worksheet.write_datetime_with_format(cell.row, cell.col, value, format)?;
                    } else {
                        worksheet.write_datetime(cell.row, cell.col, value)?;
                    }
                }
                NodeXlsxTypes::Formula((value, dynamic)) => {
                    let has_format = cell.format.is_some();
                    if dynamic && !has_format {
                        worksheet.write_dynamic_formula(cell.row, cell.col, value)?;
                    } else if dynamic && has_format {
                        let format = cell.format.unwrap();
                        let format = format_map.get(&format).unwrap();
                        worksheet
                            .write_dynamic_formula_with_format(cell.row, cell.col, value, format)?;
                    } else if !dynamic && !has_format {
                        worksheet.write_formula(cell.row, cell.col, value)?;
                    } else {
                        let format = cell.format.unwrap();
                        let format = format_map.get(&format).unwrap();
                        worksheet.write_formula_with_format(cell.row, cell.col, value, format)?;
                    }
                }
            }
        }
        return Ok(worksheet);
    }
}
