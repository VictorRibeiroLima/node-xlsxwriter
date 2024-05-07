use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNull, JsNumber, JsObject, JsString, JsUndefined, JsValue, Value},
};
use rust_xlsxwriter::Url;

pub enum NodeXlsxTypes {
    String(String),
    Number(f64),
    Link(Url),
    Unknown(String), // This is a catch-all for any type
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
            _ => {
                let js_any = any_to_string(cx, js_any)?;
                NodeXlsxTypes::Unknown(js_any)
            }
        })
    }
}

fn any_to_string<'a>(cx: &mut FunctionContext<'a>, js_any: Handle<JsValue>) -> NeonResult<String> {
    let js_string = js_any.to_string(cx)?;
    Ok(js_string.value(cx))
}

fn any_to_number<'a>(cx: &mut FunctionContext<'a>, js_any: Handle<JsValue>) -> NeonResult<f64> {
    if let Ok(num) = js_any.downcast::<JsNumber, _>(cx) {
        Ok(num.value(cx))
    } else if let Ok(str) = js_any.downcast::<JsString, _>(cx) {
        let str = str.value(cx);
        match str.parse::<f64>() {
            Ok(num) => Ok(num),
            Err(_) => {
                let error = format!("Value cannot be converted to number: {:?}", js_any);
                let js_error = cx.string(error);
                cx.throw(js_error)
            }
        }
    } else if let Ok(_) = js_any.downcast::<JsNull, _>(cx) {
        Ok(0.0)
    } else if let Ok(_) = js_any.downcast::<JsUndefined, _>(cx) {
        Ok(0.0)
    } else {
        let error = format!("Value cannot be converted to number: {:?}", js_any);
        let js_error = cx.string(error);
        cx.throw(js_error)
    }
}

fn any_to_url<'a>(cx: &mut FunctionContext<'a>, js_any: Handle<JsValue>) -> NeonResult<Url> {
    if let Ok(str) = js_any.downcast::<JsString, _>(cx) {
        return Ok(Url::new(str.value(cx)));
    }

    if let Ok(obj) = js_any.downcast::<JsObject, _>(cx) {
        let url = object_to_url(cx, obj)?;
        return Ok(url);
    }

    let error = format!("Value cannot be converted to URL: {:?}", js_any);
    let js_error = cx.string(error);
    cx.throw(js_error)
}

fn object_to_url<'a>(cx: &mut FunctionContext<'a>, obj: Handle<JsObject>) -> NeonResult<Url> {
    let url: Handle<JsString> = obj.get(cx, "url")?;
    let text: Option<Handle<JsString>> = obj.get_opt(cx, "text")?;
    let tip: Option<Handle<JsString>> = obj.get_opt(cx, "tip")?;
    let url = url.value(cx);
    let mut url = Url::new(url);
    if let Some(text) = text {
        url = url.set_text(text.value(cx));
    }
    if let Some(tip) = tip {
        url = url.set_tip(tip.value(cx));
    }
    Ok(url)
}
