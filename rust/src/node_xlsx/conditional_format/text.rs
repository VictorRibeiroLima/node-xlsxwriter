use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{ConditionalFormatText, ConditionalFormatTextRule, Format};

use crate::node_xlsx::format::NodeXlsxFormat;

use super::c_type::NodeXlsxConditionalFormatType;

pub struct Text<'a> {
    rule: ConditionalFormatTextRule,
    format: Option<&'a Format>,
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
}

impl<'a> Text<'a> {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<Self> {
        let rule: Handle<JsObject> = obj.get(cx, "rule")?;
        let rule = js_object_to_text_rule(cx, rule)?;
        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = multi_range.map(|range| range.value(cx));
        let stop_if_true: Option<Handle<JsString>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = stop_if_true.map(|stop| stop.value(cx) == "true");
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
            format,
            multi_range,
            stop_if_true,
        })
    }

    pub fn create_and_set_to_map(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        c_format_map: &mut HashMap<u32, NodeXlsxConditionalFormatType>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<u32> {
        let id: Handle<JsNumber> = obj.get(cx, "id")?;
        let id = id.value(cx) as u32;

        if !c_format_map.contains_key(&id) {
            let text = Text::from_js_object(cx, obj, format_map)?;
            let text = NodeXlsxConditionalFormatType::Text(text.into());
            c_format_map.insert(id, text);
        }
        Ok(id)
    }
}

impl Into<ConditionalFormatText> for Text<'_> {
    fn into(self) -> ConditionalFormatText {
        let mut text = ConditionalFormatText::new();
        text = text.set_rule(self.rule);
        if let Some(format) = self.format {
            text = text.set_format(format);
        }
        if let Some(multi_range) = self.multi_range {
            text = text.set_multi_range(multi_range);
        }
        if let Some(stop_if_true) = self.stop_if_true {
            text = text.set_stop_if_true(stop_if_true);
        }
        text
    }
}

fn js_object_to_text_rule(
    cx: &mut FunctionContext,
    js_object: Handle<JsObject>,
) -> NeonResult<ConditionalFormatTextRule> {
    let rule_type: Handle<JsString> = js_object.get(cx, "type")?;
    let rule_value: Handle<JsString> = js_object.get(cx, "value")?;
    let rule_type = rule_type.value(cx);
    let rule_value = rule_value.value(cx);
    match rule_type.as_str() {
        "contains" => Ok(ConditionalFormatTextRule::Contains(rule_value)),
        "doesNotContain" => Ok(ConditionalFormatTextRule::DoesNotContain(rule_value)),
        "beginsWith" => Ok(ConditionalFormatTextRule::BeginsWith(rule_value)),
        "endsWith" => Ok(ConditionalFormatTextRule::EndsWith(rule_value)),
        _ => {
            let message = format!("Unknown text rule type: {}", rule_type);
            cx.throw_error(message)
        }
    }
}
