# Word Replacement

Generate a `.docx` by using an LLM to map placeholders to your data.

## Setup

```zsh
npm install
```

## Environment

Choose **one** provider.

Create a `.env` file at the project root and set variables there.

### OpenAI

- `OPENAI_API_KEY`

### Azure OpenAI

- `AZURE_OPENAI_API_KEY`
- `AZURE_OPENAI_ENDPOINT`
- `AZURE_OPENAI_DEPLOYMENT`
- `AZURE_OPENAI_API_VERSION`

## Run example

Create a JSON file with your values, e.g. `valueMap.json`:

```json
{
  "companyName": "ABC Corporation",
  "serviceDescription": "System development and maintenance services",
  "totalAmount": "1,200,000 JPY"
}
```

Run:

```zsh
node scripts/run-example.js ./template.docx ./output.docx ./valueMap.json
```

## Smoke test (no API call)

```zsh
npm run smoke
```
