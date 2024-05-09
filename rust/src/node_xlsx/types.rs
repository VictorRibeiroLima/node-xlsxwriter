use chrono::NaiveDateTime;
use neon::{
    context::FunctionContext,
    handle::Handle,
    result::NeonResult,
    types::{JsString, JsValue},
};
use rust_xlsxwriter::{Formula, Url};

use super::util::{
    any_to_formula, any_to_naive_date_time, any_to_number, any_to_string, any_to_url,
};
pub enum NodeXlsxTypes {
    String(String),
    Number(f64),
    Link(Url),
    Date(NaiveDateTime),
    Unknown(String), // This is a catch-all for any type
    Formula((Formula, bool)),
}

impl NodeXlsxTypes {
    pub fn from_js_string<'a>(
        cx: &mut FunctionContext<'a>,
        str_type: Option<Handle<JsString>>,
        js_any: Handle<JsValue>,
    ) -> NeonResult<Self> {
        let js_string = match str_type {
            Some(str_type) => str_type,
            None => {
                let js_any = any_to_string(cx, js_any)?;
                return Ok(NodeXlsxTypes::Unknown(js_any));
            }
        };

        let js_string = js_string.value(cx);
        Ok(match js_string.to_lowercase().as_str() {
            "string" => {
                let js_any = any_to_string(cx, js_any)?;
                NodeXlsxTypes::String(js_any)
            }
            "number" => {
                let js_any = any_to_number(cx, js_any)?;
                NodeXlsxTypes::Number(js_any)
            }
            "link" => {
                let js_any = any_to_url(cx, js_any)?;
                NodeXlsxTypes::Link(js_any)
            }
            "date" => {
                let js_any = any_to_naive_date_time(cx, js_any)?;
                NodeXlsxTypes::Date(js_any)
            }
            "formula" => {
                let formula = any_to_formula(cx, js_any)?;
                NodeXlsxTypes::Formula(formula)
            }
            _ => {
                let js_any = any_to_string(cx, js_any)?;
                NodeXlsxTypes::Unknown(js_any)
            }
        })
    }
}
