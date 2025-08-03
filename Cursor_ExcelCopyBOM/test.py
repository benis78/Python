import openai

openai.api_key = "din-api-nøgle-her"

response = openai.ChatCompletion.create(
    model="gpt-4o",
    messages=[{"role": "user", "content": "Hvad er 2 + 2?"}]
)

print(response.choices[0].message["content"])
