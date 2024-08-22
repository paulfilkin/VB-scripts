# CallOpenAI_prompt Function

This VBA function `CallOpenAI_prompt` is designed to interact with the OpenAI API to generate text completions based on a user-provided message and a template. It sends a request to the OpenAI API and returns the AI-generated response.

## Functionality Overview

- **API Integration:** The function connects to the OpenAI API using a POST request, sending a user message and a system template.
- **Error Handling:** The function includes error handling to manage issues such as missing API keys, invalid worksheet references, and API request failures.
- **JSON Parsing:** The function parses the JSON response from the API to extract and return the AI's response.

## Prerequisites

- **VBA Environment:** The function is written in VBA, suitable for use in Excel or other Office applications that support VBA.
- **JSON Parser:** The function relies on a JSON parser, such as the `JsonConverter` module, which needs to be included in your VBA project. You can obtain the JSON parser [here](https://github.com/VBA-tools/VBA-JSON).
- **OpenAI API Key:** You need an API key from OpenAI, which should be stored in the specified worksheet and cell.

## Setup

1. **Worksheet Setup:**
   - Ensure you have a worksheet named `Key_models` in your workbook.
   - In cell `E3` of the `Key_models` sheet, place your OpenAI API key.
   - In cell `E5`, specify the OpenAI model you wish to use (e.g., `gpt-3.5-turbo`).

2. **Include JSON Converter:**
   - Download and import the `JsonConverter` module into your VBA project to enable JSON parsing.

## How to Use

### Function Signature

```vba
Function CallOpenAI_prompt(userMessage As String, template As String) As String
```

### Parameters

- **`userMessage`** (String): The message from the user that you want the AI to respond to.
- **`template`** (String): A system template or context provided to the AI to guide the response.

### Example Usage

```vba
Dim response As String
response = CallOpenAI_prompt("What is the weather like today?", "You are a helpful assistant.")
MsgBox response
```

In this example:
- The user message is `"What is the weather like today?"`.
- The system template is `"You are a helpful assistant."`.
- The function sends these inputs to the OpenAI API and returns the AI-generated response, which is then displayed in a message box.

### Return Value

The function returns a string containing the AI-generated response from the OpenAI API. If an error occurs, it returns an error message indicating the nature of the issue.

## Error Handling

- **Worksheet Errors:** If the `Key_models` worksheet is not found or the cells `E3` and `E5` are invalid or empty, the function will return an error message.
- **API Request Errors:** If the API request fails or if an unexpected error occurs, the function will return a detailed error message.

## Customisation

- **API Endpoint and Model:** You can customise the API endpoint and model by editing the corresponding lines in the function. Ensure the correct model is specified in the worksheet.
- **JSON Parsing:** If the API response structure changes, you may need to adjust the JSON parsing logic accordingly.

