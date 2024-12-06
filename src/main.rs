use calamine::{open_workbook, DataType, Reader, Xlsx};
use http::header::HeaderMap;
use reqwest::Client;
// use std::collections::HashMap;

#[tokio::main]
async fn main() -> Result<(), reqwest::Error> {
    let client = Client::new();

    let mut headers = HeaderMap::new();

    headers.insert("api-key", "test_admin".parse().unwrap());

    // Change the excel file name
    let mut workbook: Xlsx<_> =
        open_workbook("/Users/chethan.rao/playground/GSM_script/src/error_category.xlsx")
            .expect("Cannot open file");

    let mut count = 1;
    // Change the sheet name
    if let Some(Ok(r)) = workbook.worksheet_range("adyen") {
        for (ind, row) in r.rows().enumerate() {
            // delay for 1 second
            tokio::time::sleep(std::time::Duration::from_millis(100)).await;
            if ind == 0 {
                continue;
            }

            let connector = row[0].get_string();
            let code = match &row[2] {
                DataType::Int(i) => Some(i.to_string()),
                DataType::String(s) => Some(s.clone()),
                DataType::Float(f) => Some(f.to_string()),
                _ => None,
            };
            let message = match &row[3] {
                DataType::Int(i) => Some(i.to_string()),
                DataType::String(s) => Some(s.clone()),
                DataType::Float(f) => Some(f.to_string()),
                _ => None,
            };
            let error_category = row[4].get_string().map(|inner| inner.trim());
            let decision = Some("do_default");

            // log the values
            println!(
                "{}. connector: {:?}, code: {:?}, message: {:?}, error_category: {:?}\n",
                count, connector, code, message, error_category
            );

            match (connector, code, message, decision, error_category) {
                (
                    Some(connector),
                    Some(code),
                    Some(message),
                    Some(decision),
                    Some(error_category),
                ) => {
                    let insert_body = serde_json::json!({
                        "connector": connector,
                        "flow": "Authorize",
                        "sub_flow": "sub_flow",
                        "code": code,
                        "message": message,
                        "status": "Failure",
                        "decision": decision,
                        "step_up_possible": false,
                        "error_category": error_category
                    });

                    let update_body = serde_json::json!({
                        "connector": connector,
                        "flow": "Authorize",
                        "sub_flow": "sub_flow",
                        "code": code,
                        "message": message,
                        "error_category": error_category
                    });
                    insert_gsm(
                        client.to_owned(),
                        insert_body,
                        update_body,
                        headers.to_owned(),
                    )
                    .await?;
                }
                _ => print!("Message or code is empty"),
            }
            count += 1;

            println!("\n||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||\n");
        }
    }
    Ok(())
}

async fn insert_gsm(
    client: Client,
    insert_body: serde_json::Value,
    update_body: serde_json::Value,
    header: HeaderMap,
) -> Result<(), reqwest::Error> {
    let response = client
        // .post("http://localhost:8080/gsm")
        .post("https://sandbox.hyperswitch.io/gsm")
        // .post("https://webhook.site/3762a5ec-0f8e-4dc1-8fe3-1adf6ca64d40/gsm")
        .headers(header.clone())
        .json(&insert_body)
        .send()
        .await?;

    let resp = response.text().await.unwrap();
    println!("insert response - {:?}", resp);

    if resp.contains("GSM with given key already exists in our records") {
        update_gsm(client, update_body, header).await?;
    }

    Ok(())
}

async fn update_gsm(
    client: Client,
    update_body: serde_json::Value,
    header: HeaderMap,
) -> Result<(), reqwest::Error> {
    let response = client
        // .post("http://localhost:8080/gsm/update")
        .post("https://sandbox.hyperswitch.io/gsm")
        // .post("https://webhook.site/3762a5ec-0f8e-4dc1-8fe3-1adf6ca64d40/gsm")
        .headers(header)
        .json(&update_body)
        .send()
        .await?;

    let resp = response.text().await.unwrap();
    println!("update response - {:?}", resp);

    Ok(())
}
