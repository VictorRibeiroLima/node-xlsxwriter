use neon::prelude::*;
use neon::types::buffer::TypedArray;
use node_xlsx::NodeXlsxWorkbook;

mod node_xlsx;

fn hello(mut cx: FunctionContext) -> JsResult<JsString> {
    Ok(cx.string("hello node"))
}

fn save_to_buffer(mut cx: FunctionContext) -> JsResult<JsPromise> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let promise = cx
        .task(move || workbook.save_to_buffer())
        .promise(move |mut cx, greeting| match greeting {
            Ok(buffer) => {
                let mut js_buffer = cx.buffer(buffer.len())?;
                for (i, elem) in js_buffer.as_mut_slice(&mut cx).iter_mut().enumerate() {
                    *elem = buffer[i];
                }
                Ok(js_buffer)
            }
            Err(err) => cx.throw_error(err.to_string()),
        });

    Ok(promise)
}

fn save_to_buffer_sync(mut cx: FunctionContext) -> JsResult<JsBuffer> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let buffer = workbook.save_to_buffer().unwrap();

    let mut js_buffer = cx.buffer(buffer.len())?;
    for (i, elem) in js_buffer.as_mut_slice(&mut cx).iter_mut().enumerate() {
        *elem = buffer[i];
    }

    Ok(js_buffer)
}

fn save_to_file_sync(mut cx: FunctionContext) -> JsResult<JsUndefined> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let path: Handle<JsString> = cx.argument(1)?;
    let path = path.value(&mut cx);

    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let _ = workbook.save_to_file(&path).unwrap();

    Ok(cx.undefined())
}

#[neon::main]
fn main(mut cx: ModuleContext) -> NeonResult<()> {
    cx.export_function("hello", hello)?;
    cx.export_function("saveToBuffer", save_to_buffer)?;
    cx.export_function("saveToBufferSync", save_to_buffer_sync)?;
    cx.export_function("saveToFileSync", save_to_file_sync)?;
    Ok(())
}
