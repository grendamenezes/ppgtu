import docx
from docxtpl import DocxTemplate

doc = DocxTemplate("modelo_atividades.docx")

nomes = ["João", "Pedro", "Jorge"]
dados = {
    "programa": "\n".join(nomes)
}
doc.render(dados)
doc.save("meu_arquivo.docx")
