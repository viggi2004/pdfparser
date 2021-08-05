from py_pdf_parser.loaders import load_file
from py_pdf_parser.visualise import visualise
import tablib

def extract(elements, data, position):
    for elem in elements:
        currentContact = ()
        nameText = (document.elements.after(elem)[0]).text()
        typeText = (document.elements.after(elem)[1]).text()
        genderText = (document.elements.after(elem)[2]).text()
        emailText = (document.elements.after(elem)[3]).text()
        ContactText = (document.elements.after(elem)[4]).text()
        nameValue = (document.elements.after(elem)[5]).text()
        typeValue = (document.elements.after(elem)[6]).text()
        genderValue = (document.elements.after(elem)[7]).text()
        emailValue = (document.elements.after(elem)[8]).text()
        contactValue = (document.elements.after(elem)[9]).text()
        currentContact = (nameValue, typeValue, genderValue, emailValue, contactValue, position)
        data.append(currentContact)
    with open('./output/first.xlsx', 'wb') as f:
        f.write(data.export('xlsx'))

document = load_file("./input/report-cut.pdf")
principal_innovator_elements = document.elements.filter_by_text_equal("Principal Innovator")#.extract_single_element()
co_innovator_elements = document.elements.filter_by_text_equal("Co-Innovator")#.extract_single_element()

contacts = []
data = tablib.Dataset(headers=['Name', 'Type', 'Gender','Email','Contact', 'Position'])
extract(principal_innovator_elements, data, 'Principal Innovator')
extract(co_innovator_elements, data, 'Co-Innovator')
