use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{Format, Formula, TableColumn, TableFunction};

use crate::node_xlsx::{format::NodeXlsxFormat, util::object_to_formula};

use super::util::object_to_table_function;

pub struct NodeXlsxTableColumn<'a> {
    formula: Option<Formula>,
    format: Option<&'a Format>,
    header: Option<String>,
    header_format: Option<&'a Format>,
    total_label: Option<String>,
    total_function: Option<TableFunction>,
}

impl<'a> NodeXlsxTableColumn<'a> {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<Self> {
        let mut format_id = None;
        let formula: Option<Handle<JsObject>> = obj.get_opt(cx, "formula")?;
        let formula = match formula {
            Some(formula) => {
                let (formula, _) = object_to_formula(cx, formula)?;
                Some(formula)
            }
            None => None,
        };

        let format: Option<Handle<JsObject>> = obj.get_opt(cx, "format")?;
        match format {
            Some(format) => {
                let id: Handle<JsNumber> = format.get(cx, "id")?;
                let id = id.value(cx) as u32;
                if !format_map.contains_key(&id) {
                    let format = NodeXlsxFormat::from_js_object(cx, format)?;
                    format_map.insert(id, format.into());
                }
                format_id = Some(id);
            }
            None => {}
        };

        let header_format: Option<Handle<JsObject>> = obj.get_opt(cx, "headerFormat")?;
        let header_format = match header_format {
            Some(header_format) => {
                let id: Handle<JsNumber> = header_format.get(cx, "id")?;
                let id = id.value(cx) as u32;
                if !format_map.contains_key(&id) {
                    let format = NodeXlsxFormat::from_js_object(cx, header_format)?;
                    format_map.insert(id, format.into());
                }
                let format = format_map.get(&id).unwrap();
                Some(format)
            }
            None => None,
        };
        let format = format_id.and_then(|id| format_map.get(&id));

        let header: Option<Handle<JsString>> = obj.get_opt(cx, "header")?;
        let header = header.map(|header| header.value(cx));

        let total_label: Option<Handle<JsString>> = obj.get_opt(cx, "totalLabel")?;
        let total_label = total_label.map(|total_label| total_label.value(cx));

        let total_function: Option<Handle<JsObject>> = obj.get_opt(cx, "totalFunction")?;
        let total_function = match total_function {
            Some(total_function) => {
                let total_function = object_to_table_function(cx, total_function)?;
                Some(total_function)
            }
            None => None,
        };

        Ok(Self {
            formula,
            format,
            header,
            header_format,
            total_label,
            total_function,
        })
    }

    pub fn create_and_into(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<TableColumn> {
        let column = Self::from_js_object(cx, obj, format_map)?;
        Ok(column.into())
    }
}

impl Into<TableColumn> for NodeXlsxTableColumn<'_> {
    fn into(self) -> TableColumn {
        let mut column = TableColumn::new();
        if let Some(formula) = self.formula {
            column = column.set_formula(formula);
        }
        if let Some(format) = self.format {
            column = column.set_format(format);
        }
        if let Some(header) = self.header {
            column = column.set_header(header);
        }
        if let Some(header_format) = self.header_format {
            column = column.set_header_format(header_format);
        }
        if let Some(total_label) = self.total_label {
            column = column.set_total_label(total_label);
        }
        if let Some(total_function) = self.total_function {
            column = column.set_total_function(total_function);
        }
        column
    }
}
