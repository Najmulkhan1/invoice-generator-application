from docxtpl import DocxTemplate

# Create a document object
doc = DocxTemplate("yyyyupdate.docx")


invoice_list = [[2, "pen", 0.5, 1],
                [1, "Paper pack", 4, 7],
                [3, "notebook", 7, 4]]



# Define the context for rendering the template

#context = {
    #'name': 'John Doe',
    #'Phone': '017657777234',
    #'invoice_list': invoice_list
#}

context = {
    'name': "mukit",
    'phone': "01767382767",
    'email': "najmulislam732",
    'address': "busdd",
    'invoice': "8787878",
    'invoice_list': invoice_list
}


# Render the document with the context
doc.render(context)

# Save the generated document
doc.save("generatedqqmmm.docx")
