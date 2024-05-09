use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{ConditionalFormatError, Format};

use crate::node_xlsx::format::NodeXlsxFormat;

use super::c_type::NodeXlsxConditionalFormatType;

pub struct Error<'a> {
    invert: bool,
    format: Option<&'a Format>,
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
}

impl<'a> Error<'a> {
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
            let error = Error::from_js_object(cx, obj, format_map)?;
            c_format_map.insert(id, NodeXlsxConditionalFormatType::Error(error.into()));
        }
        Ok(id)
    }
}

impl Into<ConditionalFormatError> for Error<'_> {
    fn into(self) -> ConditionalFormatError {
        let mut error = ConditionalFormatError::new();
        if self.invert {
            error = error.invert();
        }
        if let Some(format) = self.format {
            error = error.set_format(format);
        }
        if let Some(multi_range) = self.multi_range {
            error = error.set_multi_range(&multi_range);
        }
        if let Some(stop_if_true) = self.stop_if_true {
            error = error.set_stop_if_true(stop_if_true);
        }
        error
    }
}
