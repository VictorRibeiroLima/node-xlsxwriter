use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNumber, JsObject, JsValue},
};

#[derive(Clone, Copy)]
pub struct Color {
    red: u8,
    green: u8,
    blue: u8,
}

impl Color {
    pub fn from_js_object(cx: &mut FunctionContext, object: Handle<JsObject>) -> NeonResult<Self> {
        let red: Handle<JsValue> = object.get(cx, "red")?;
        let green: Handle<JsValue> = object.get(cx, "green")?;
        let blue: Handle<JsValue> = object.get(cx, "blue")?;
        Ok(Self {
            red: red.downcast_or_throw::<JsNumber, _>(cx)?.value(cx) as u8,
            green: green.downcast_or_throw::<JsNumber, _>(cx)?.value(cx) as u8,
            blue: blue.downcast_or_throw::<JsNumber, _>(cx)?.value(cx) as u8,
        })
    }
}

impl Into<rust_xlsxwriter::Color> for Color {
    fn into(self) -> rust_xlsxwriter::Color {
        let red = self.red as u32;
        let green = self.green as u32;
        let blue = self.blue as u32;
        rust_xlsxwriter::Color::RGB((red << 16) | (green << 8) | blue)
    }
}
