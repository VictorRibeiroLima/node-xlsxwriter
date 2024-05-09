use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsDate, JsNumber, JsObject, JsString, JsValue},
};
use rust_xlsxwriter::{
    ConditionalFormatCell, ConditionalFormatCellRule, Format, IntoConditionalFormatValue,
};

use crate::node_xlsx::{format::NodeXlsxFormat, util::js_date_to_naive_date_time};

use super::{c_type::NodeXlsxConditionalFormatType, rule::NodeXlsxConditionalFormatCellRule};

pub struct Cell<'a> {
    rule: Option<NodeXlsxConditionalFormatCellRule<'a>>,
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
    format: Option<&'a Format>,
}

impl<'a> Cell<'a> {
    pub fn from_js_object(
        cx: &'a mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &'a mut HashMap<u32, Format>,
        rule: Option<NodeXlsxConditionalFormatCellRule<'a>>,
    ) -> NeonResult<Self> {
        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = multi_range.map(|range| range.value(cx));

        let stop_if_true: Option<Handle<JsBoolean>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = stop_if_true.map(|stop| stop.value(cx));

        let format: Option<Handle<JsObject>> = obj.get_opt(cx, "format")?;
        let format = match format {
            Some(format) => {
                let id: Handle<JsNumber> = format.get(cx, "id")?;
                let id = id.value(cx) as u32;
                if !format_map.contains_key(&id) {
                    let format = NodeXlsxFormat::from_js_object(cx, format)?;
                    format_map.insert(id, format.into());
                }
                let format = format_map.get(&id).unwrap();
                Some(format)
            }
            None => None,
        };

        Ok(Self {
            rule,
            multi_range,
            stop_if_true,
            format,
        })
    }

    //TODO: rewrite this function for better readability
    pub fn create_and_set_to_map(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        c_format_map: &mut HashMap<u32, NodeXlsxConditionalFormatType>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<u32> {
        let id: Handle<JsNumber> = obj.get(cx, "id")?;
        let id = id.value(cx) as u32;

        //Scope hack
        let naive;
        let optional_naive;

        let rule: Option<Handle<JsObject>> = obj.get_opt(cx, "rule")?;
        let rule = if let Some(rule) = rule {
            let r_type: Handle<JsString> = rule.get(cx, "type")?;
            let value: Handle<JsValue> = rule.get(cx, "value")?;
            let optional_value: Option<Handle<JsValue>> = rule.get_opt(cx, "optionalValue")?;
            let r_type = r_type.value(cx);
            let value = if let Ok(string) = value.downcast::<JsString, _>(cx) {
                let string = string.value(cx);
                let optional_value = match optional_value {
                    Some(optional_value) => {
                        let optional_value = optional_value.downcast_or_throw::<JsString, _>(cx)?;
                        Some(optional_value.value(cx))
                    }
                    None => None,
                };
                let variant = set_value_to_writer_variant(cx, &r_type, string, optional_value)?;
                let variant = NodeXlsxConditionalFormatCellRule::String(variant);
                Some(variant)
            } else if let Ok(number) = value.downcast::<JsNumber, _>(cx) {
                let number = number.value(cx);
                let optional_value = match optional_value {
                    Some(optional_value) => {
                        let optional_value = optional_value.downcast_or_throw::<JsNumber, _>(cx)?;
                        Some(optional_value.value(cx))
                    }
                    None => None,
                };
                let variant = set_value_to_writer_variant(cx, &r_type, number, optional_value)?;
                let variant = NodeXlsxConditionalFormatCellRule::Number(variant);
                Some(variant)
            } else if let Ok(date) = value.downcast::<JsDate, _>(cx) {
                naive = Some(js_date_to_naive_date_time(cx, date)?);
                let naive_ref = naive.as_ref().unwrap();
                optional_naive = match optional_value {
                    Some(optional_value) => {
                        let optional_value = optional_value.downcast_or_throw::<JsDate, _>(cx)?;
                        let optional_value = js_date_to_naive_date_time(cx, optional_value)?;
                        Some(optional_value)
                    }
                    None => None,
                };
                let optional_naive_ref = optional_naive.as_ref();
                let variant =
                    set_value_to_writer_variant(cx, &r_type, naive_ref, optional_naive_ref)?;
                Some(NodeXlsxConditionalFormatCellRule::DateTime(variant))
            } else {
                None
            };

            value
        } else {
            None
        };

        let cell = Cell::from_js_object(cx, obj, format_map, rule)?;
        let cell = cell.into();
        let c_format = NodeXlsxConditionalFormatType::Cell(cell);
        c_format_map.insert(id, c_format);
        Ok(id)
    }
}

impl Into<ConditionalFormatCell> for Cell<'_> {
    fn into(self) -> ConditionalFormatCell {
        let mut cell = ConditionalFormatCell::new();
        if let Some(rule) = self.rule {
            match rule {
                NodeXlsxConditionalFormatCellRule::String(rule) => {
                    cell = cell.set_rule(rule);
                }
                NodeXlsxConditionalFormatCellRule::Number(rule) => {
                    cell = cell.set_rule(rule);
                }
                NodeXlsxConditionalFormatCellRule::DateTime(rule) => {
                    cell = cell.set_rule(rule);
                }
            }
        }

        if let Some(multi_range) = self.multi_range {
            cell = cell.set_multi_range(multi_range);
        }

        if let Some(stop_if_true) = self.stop_if_true {
            cell = cell.set_stop_if_true(stop_if_true);
        }

        if let Some(format) = self.format {
            cell = cell.set_format(format);
        }

        cell
    }
}

fn set_value_to_writer_variant<T: IntoConditionalFormatValue>(
    cx: &mut FunctionContext,
    r_type: &str,
    value: T,
    optional_value: Option<T>,
) -> NeonResult<ConditionalFormatCellRule<T>> {
    match r_type {
        "equalTo" => Ok(ConditionalFormatCellRule::EqualTo(value)),
        "notEqualTo" => Ok(ConditionalFormatCellRule::NotEqualTo(value)),
        "greaterThan" => Ok(ConditionalFormatCellRule::GreaterThan(value)),
        "greaterThanOrEqualTo" => Ok(ConditionalFormatCellRule::GreaterThanOrEqualTo(value)),
        "lessThan" => Ok(ConditionalFormatCellRule::LessThan(value)),
        "lessThanOrEqualTo" => Ok(ConditionalFormatCellRule::LessThanOrEqualTo(value)),
        "between" => match optional_value {
            Some(optional_value) => Ok(ConditionalFormatCellRule::Between(value, optional_value)),
            None => cx.throw_error("Missing second value for 'between' rule"),
        },
        "notBetween" => match optional_value {
            Some(optional_value) => {
                Ok(ConditionalFormatCellRule::NotBetween(value, optional_value))
            }
            None => cx.throw_error("Missing second value for 'notBetween' rule"),
        },
        _ => {
            let err = format!("Invalid ConditionalFormatType: {}", r_type);
            cx.throw_error(err)
        }
    }
}
