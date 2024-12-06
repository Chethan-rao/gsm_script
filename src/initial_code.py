use calamine::{open_workbook, Reader, Xlsx};
use http::header::HeaderMap;
use reqwest::Client;
use std::collections::HashMap;

#[tokio::main]
async fn main() -> Result<(), reqwest::Error> {
    let client = Client::new();

    let mut headers = HeaderMap::new();
    headers.insert("api-key", "test_admin".parse().unwrap());

    # Change the excel file name
    let mut workbook: Xlsx<_> = open_workbook(
        "/Users/chethan.rao/rust_practise/gsm_script/src/Stripe_redirect_errors.xlsx",
    )
    .expect("Cannot open file");

    # Change the sheet name
    if let Some(Ok(r)) = workbook.worksheet_range("Sheet 1 - Stripe Redirect error") {
        for (ind, row) in r.rows().enumerate() {
            if ind == 0 || ind == 1 {
                continue;
            }

            let code = row[0].get_string();
            let message = row[1].get_string();
            if let Some(code) = code {
                let mut body = HashMap::new();

                body.insert("connector", "stripe");
                body.insert("flow", "flow3");
                body.insert("sub_flow", "sub_flow");
                body.insert("code", code);
                body.insert("status", "status1");
                body.insert("decision", "retry");
                body.insert("message", message.unwrap_or(""));

                add_gsm(client.to_owned(), body, headers.to_owned()).await?;
            }
        }
    }

    Ok(())
}

async fn add_gsm(
    client: Client,
    body: HashMap<&str, &str>,
    header: HeaderMap,
) -> Result<(), reqwest::Error> {
    let response = client
        .post("http:ocalhost:8080/gsm")
        .headers(header)
        .json(&body)
        .send()
        .await?;

    println!("response {}", response.text().await?);
    Ok(())
}
