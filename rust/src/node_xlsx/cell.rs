use super::types::NodeXlsxTypes;
use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNull, JsNumber, JsObject, JsString, JsUndefined, JsValue},
};

#[derive(Debug)]
pub struct NodeXlsxCell {
    pub col: u16,
    pub row: u32,
    pub value: String,
    pub cell_type: NodeXlsxTypes,
}

impl NodeXlsxCell {
    pub fn from_js_object<'a>(
        cx: &mut FunctionContext<'a>,
        obj: Handle<JsObject>,
    ) -> NeonResult<Self> {
        let col: Handle<JsNumber> = obj.get(cx, "col")?;
        let col = col.value(cx);
        if col < 0.0 || col >= 16_384.0 {
            let error = format!("Column with illegal number {}", col.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let col = u16::try_from(col as u64);
        let col = match col {
            Ok(col) => col,
            Err(_) => {
                let js_string = cx.string("Column number is too large");
                return cx.throw(js_string);
            }
        };

        let row: Handle<JsNumber> = obj.get(cx, "row")?;
        let row = row.value(cx);
        if row < 0.0 || row >= 1_048_577.0 {
            let error = format!("Row with illegal number {}", row.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let row = row as u32;

        let cel_type: Handle<JsString> = obj.get(cx, "cellType")?;
        let cel_type = NodeXlsxTypes::from_js_string(cx, cel_type)?;

        let value: Handle<JsValue> = obj.get(cx, "value")?;
        let value = if let Ok(str) = value.downcast::<JsString, _>(cx) {
            str.value(cx)
        } else if let Ok(num) = value.downcast::<JsNumber, _>(cx) {
            num.value(cx).to_string()
        } else if let Ok(_) = value.downcast::<JsNull, _>(cx) {
            "".to_string()
        } else if let Ok(_) = value.downcast::<JsUndefined, _>(cx) {
            "".to_string()
        } else {
            return cx.throw_type_error("expected string or number");
        };

        Ok(Self {
            col,
            row,
            value,
            cell_type: cel_type,
        })
    }
}
