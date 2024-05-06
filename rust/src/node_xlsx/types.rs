use neon::{context::FunctionContext, handle::Handle, result::NeonResult, types::JsString};

#[derive(Debug)]
pub enum NodeXlsxTypes {
    String,
    Number,
    Link,
}

impl NodeXlsxTypes {
    pub fn from_js_string<'a>(
        cx: &mut FunctionContext<'a>,
        js_string: Handle<JsString>,
    ) -> NeonResult<Self> {
        let js_string = js_string.value(cx);
        Ok(match js_string.to_lowercase().as_str() {
            "string" => NodeXlsxTypes::String,
            "number" => NodeXlsxTypes::Number,
            "link" => NodeXlsxTypes::Link,
            _ => panic!("Invalid cell type"),
        })
    }
}
