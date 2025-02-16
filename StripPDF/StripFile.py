import textract
fn="star_ef_produktkatalog_09-2018_komprimiert"
text = textract.process("C:/Users/Jesper/Downloads/star_ef_produktkatalog_09-2018_komprimiert.pdf")

print(text)
