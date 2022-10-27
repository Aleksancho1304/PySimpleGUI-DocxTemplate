import PySimpleGUI as sg
from docxtpl import DocxTemplate
i=0
layout = [[sg.Text('Сотрудник 1'),sg.Checkbox('Присутствие',enable_events=True,key = ('Yes1')),sg.Text('Причина отсутствия: '), sg.InputText()],
          [sg.Text('Сотрудник 2'),sg.Checkbox('Присутствие',enable_events=True,key = ('Yes2')),sg.Text('Причина отсутствия: '), sg.InputText()],
          [sg.Button('Загрузить',enable_events=True, key='Zagruzit', font='Helvetica 14')]]
#layout = [{sg.Text('сотрудник 1'):sg.Checkbox('Присутствие',default=False)}]

window = sg.Window('Список сотрудников', layout)

#b = layout[0][0].DisplayText
#словарь {.docx метка сотрудника:имя сотрудника из приложения}
context = {f'sotrudnik{i+1}': '' }

def func():
 doc = DocxTemplate("document.docx")
 doc.render(context)
 doc.save("director-final.docx")

while True:                             # The Event Loop
    event, values = window.read()
    # print(event, values) #debug
    #print(a)
    if event in (None, 'Exit', 'Cancel'):
        break
   # if event == ('Yes',1):
    #    print(values[('Yes',1)])
    if event == ('Zagruzit'):
        for i in range(2):
            if values[(f'Yes{i+1}')] == True:
               context[f'sotrudnik{i+1}']= layout[i][0].DisplayText

           # else:
           #    break #заменяем значение на '_' пустоту
    #print(context)
    func()
window.close()
