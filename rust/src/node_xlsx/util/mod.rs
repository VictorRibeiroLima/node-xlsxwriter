use chrono::{DateTime, NaiveDateTime};
use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsDate, JsNull, JsNumber, JsObject, JsString, JsUndefined, JsValue, Value},
};
use rust_xlsxwriter::{Formula, Url};

pub fn any_to_string<'a>(
    cx: &mut FunctionContext<'a>,
    js_any: Handle<JsValue>,
) -> NeonResult<String> {
    if let Ok(_) = js_any.downcast::<JsNull, _>(cx) {
        return Ok("".to_string());
    }
    if let Ok(_) = js_any.downcast::<JsUndefined, _>(cx) {
        return Ok("".to_string());
    }
    let js_string = js_any.to_string(cx)?;
    Ok(js_string.value(cx))
}

pub fn any_to_number<'a>(cx: &mut FunctionContext<'a>, js_any: Handle<JsValue>) -> NeonResult<f64> {
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

pub fn any_to_url<'a>(cx: &mut FunctionContext<'a>, js_any: Handle<JsValue>) -> NeonResult<Url> {
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

pub fn any_to_naive_date_time<'a>(
    cx: &mut FunctionContext<'a>,
    js_any: Handle<JsValue>,
) -> NeonResult<NaiveDateTime> {
    let js_date = js_any.downcast_or_throw::<JsDate, _>(cx)?;
    js_date_to_naive_date_time(cx, js_date)
}

pub fn js_date_to_naive_date_time<'a>(
    cx: &mut FunctionContext<'a>,
    js_date: Handle<JsDate>,
) -> NeonResult<NaiveDateTime> {
    let timestamp = js_date.value(cx);
    let naive_date_time = DateTime::from_timestamp_millis(timestamp as i64);
    match naive_date_time {
        Some(naive_date_time) => Ok(naive_date_time.naive_utc()),
        None => {
            Ok(NaiveDateTime::parse_from_str("1970-01-01 00:00:00", "%Y-%m-%d %H:%M:%S").unwrap())
        }
    }
}

pub fn any_to_formula<'a>(
    cx: &mut FunctionContext<'a>,
    js_any: Handle<JsValue>,
) -> NeonResult<Formula> {
    if let Ok(obj) = js_any.downcast::<JsObject, _>(cx) {
        let formula = object_to_formula(cx, obj)?;
        return Ok(formula);
    }

    let error = format!("Value cannot be converted to formula: {:?}", js_any);
    let js_error = cx.string(error);
    cx.throw(js_error)
}

pub fn object_to_formula<'a>(
    cx: &mut FunctionContext<'a>,
    obj: Handle<JsObject>,
) -> NeonResult<rust_xlsxwriter::Formula> {
    let formula: Handle<JsString> = obj.get(cx, "formula")?;
    let formula = formula.value(cx);

    let result: Option<Handle<JsString>> = obj.get_opt(cx, "result")?;
    let result = match result {
        Some(result) => Some(result.value(cx)),
        None => None,
    };

    let use_future_functions: Option<Handle<JsBoolean>> = obj.get_opt(cx, "useFutureFunctions")?;
    let use_future_functions = match use_future_functions {
        Some(use_future_functions) => Some(use_future_functions.value(cx)),
        None => None,
    };

    let use_table_functions: Option<Handle<JsBoolean>> = obj.get_opt(cx, "useTableFunctions")?;
    let use_table_functions = match use_table_functions {
        Some(use_table_functions) => Some(use_table_functions.value(cx)),
        None => None,
    };

    let mut formula = Formula::new(formula);

    if let Some(result) = result {
        formula = formula.set_result(result);
    }

    if let Some(use_future_functions) = use_future_functions {
        if use_future_functions {
            formula = formula.use_future_functions()
        }
    }

    if let Some(use_table_functions) = use_table_functions {
        if use_table_functions {
            formula = formula.use_table_functions()
        }
    }

    Ok(formula)
}

pub fn object_to_url<'a>(cx: &mut FunctionContext<'a>, obj: Handle<JsObject>) -> NeonResult<Url> {
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
