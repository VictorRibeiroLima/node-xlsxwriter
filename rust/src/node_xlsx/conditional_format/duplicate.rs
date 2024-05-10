use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{ConditionalFormatDuplicate, Format};

use crate::node_xlsx::format::NodeXlsxFormat;

use super::c_type::NodeXlsxConditionalFormatType;

pub struct Duplicate<'a> {
    invert: bool,
    format: Option<&'a Format>,
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
}

impl<'a> Duplicate<'a> {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<Self> {
        let invert: Handle<JsBoolean> = obj.get(cx, "invert")?;
        let invert = invert.value(cx);

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

        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = multi_range.map(|range| range.value(cx));

        let stop_if_true: Option<Handle<JsBoolean>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = stop_if_true.map(|stop| stop.value(cx));

        Ok(Self {
            invert,
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
            let duplicate = Duplicate::from_js_object(cx, obj, format_map)?;
            c_format_map.insert(
                id,
                NodeXlsxConditionalFormatType::Duplicate(duplicate.into()),
            );
        }
        Ok(id)
    }
}

impl Into<ConditionalFormatDuplicate> for Duplicate<'_> {
    fn into(self) -> ConditionalFormatDuplicate {
        let mut duplicate = ConditionalFormatDuplicate::new();
        if self.invert {
            duplicate = duplicate.invert();
        }
        if let Some(format) = self.format {
            duplicate = duplicate.set_format(format);
        }
        if let Some(multi_range) = self.multi_range {
            duplicate = duplicate.set_multi_range(&multi_range);
        }
        if let Some(stop_if_true) = self.stop_if_true {
            duplicate = duplicate.set_stop_if_true(stop_if_true);
        }
        duplicate
    }
}
