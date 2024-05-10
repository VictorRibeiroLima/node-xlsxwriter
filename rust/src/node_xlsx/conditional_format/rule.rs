use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsDate, JsNumber, JsObject, JsString, JsValue},
};
use rust_xlsxwriter::{ConditionalFormatCellRule, ConditionalFormatType, ConditionalFormatValue};

use crate::node_xlsx::util::js_date_to_naive_date_time;

pub struct NodeXlsxFormatTypeRule {
    pub r_type: ConditionalFormatType,
    pub value: ConditionalFormatValue,
}

impl NodeXlsxFormatTypeRule {
    pub fn from_js_value(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let r_type: Handle<JsString> = obj.get(cx, "type")?;
        let r_type = type_from_js_string(cx, r_type)?;
        let value: Handle<JsValue> = obj.get(cx, "value")?;
        let value = value_from_js_value(cx, value)?;
        Ok(Self { r_type, value })
    }
}

pub enum NodeXlsxConditionalFormatCellRule<'a> {
    String(ConditionalFormatCellRule<String>),
    Number(ConditionalFormatCellRule<f64>),
    DateTime(ConditionalFormatCellRule<&'a chrono::NaiveDateTime>),
}

pub fn type_from_js_string(
    cx: &mut FunctionContext,
    value: Handle<JsString>,
) -> NeonResult<ConditionalFormatType> {
    let value = value.value(cx);
    match value.as_str() {
        "number" => Ok(ConditionalFormatType::Number),
        "percent" => Ok(ConditionalFormatType::Percent),
        "formula" => Ok(ConditionalFormatType::Formula),
        "percentile" => Ok(ConditionalFormatType::Percentile),
        "automatic" => Ok(ConditionalFormatType::Automatic),
        "lowest" => Ok(ConditionalFormatType::Lowest),
        "highest" => Ok(ConditionalFormatType::Highest),
        _ => {
            let err = format!("Invalid ConditionalFormatType: {}", value);
            cx.throw_error(err)
        }
    }
}

fn value_from_js_value(
    cx: &mut FunctionContext,
    value: Handle<JsValue>,
) -> NeonResult<ConditionalFormatValue> {
    if let Ok(string) = value.downcast::<JsString, _>(cx) {
        let string = string.value(cx);
        return Ok(string.into());
    } else if let Ok(number) = value.downcast::<JsNumber, _>(cx) {
        let number = number.value(cx);
        return Ok(number.into());
    } else if let Ok(date) = value.downcast::<JsDate, _>(cx) {
        let naive = &js_date_to_naive_date_time(cx, date)?;
        return Ok(naive.into());
    }
    let err = format!("Invalid ConditionalFormatValue: {:?}", value);
    cx.throw_error(err)
}
