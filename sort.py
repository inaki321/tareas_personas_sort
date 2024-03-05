import pandas as pd
import random
import xlrd
from xlutils.copy import copy


#inputFile = pd.read_excel('principal_personas_tareas.xls', sheet_name='inputs_nombres_areas')
inputFile = pd.read_excel('input_personas_tareas.xlsx')
inputAsigned = pd.read_excel('historial_noTouch.xlsx')

#inputAsigned = pd.read_excel('historial.xls', sheet_name='historial')

personsDf = inputFile['Persona']
dutiesDf = inputFile['areas']

"""
OBTENER PERSONAS Y TAREAS POR SEPARADO
"""
for i in range(len(personsDf)):
    if personsDf.iloc[i] not in inputAsigned['Persona_asignada'].values:
        newRow = {
            'Persona': personsDf.iloc[i], 'Persona_asignada': '', 'Tarea_asignada': ''}
        inputAsigned = inputAsigned._append(newRow, ignore_index=True)

# sort for more randomness
inputAsigned = inputAsigned.sample(frac=1).reset_index(drop=True)


"""
GENERACION DE PERSONAS
"""
# need to remove from the df the already used duties and persons
personsDf = personsDf.dropna()
personsList = personsDf.values.tolist()


def searchForPrsons(currentPerson, assigned):
    if(isinstance(assigned, float)):
        assigned = ''
    assigned = assigned.split('||')
    assignedPerson = assigned[-1]
    possibleError = False
    # empty assigned
    if (assignedPerson == ''):
        if (len(personsList) == 1 and personsList[0] == currentPerson):
            possibleError = True
        if (not possibleError):
            randomPerson = random.choice(
                [personidx for personidx in personsList if personidx != currentPerson])
            assigned.pop(0)
    # if assigned person has already been asigned the last month, then
    else:
        if (len(personsList) == 1 and personsList[0] == currentPerson):
            possibleError = True
        if (sorted(assigned) == sorted(personsList)):
            possibleError = True

        if (not possibleError):
            randomPerson = random.choice(
                [personidx for personidx in personsList if personidx != currentPerson])
            while randomPerson in assigned:
                randomPerson = random.choice(
                    [personidx for personidx in personsList if personidx != currentPerson])

    if (possibleError):
        return False

    assigned.append(randomPerson)
    if (randomPerson in personsList):
        personsList.remove(randomPerson)

    return assigned


# loop for assigned persons

assignedOcured = []
for idx in range(len(inputAsigned)):
    idxrow = inputAsigned.iloc[idx]
    assignedArray = []
    if (len(personsList) > 0):
        assignedArray = searchForPrsons(
            idxrow['Persona'], idxrow['Persona_asignada'])

    if (assignedArray):
        joinedArray = "||".join(assignedArray)
        inputAsigned.iloc[idx]['Persona_asignada'] = joinedArray
        assignedOcured.append(True)
    else:
        assignedOcured.append(False)
print('PERSONAS GENERADAS CORRECTO ')
inputAsigned['assigned_ocured_person'] = assignedOcured



"""
GENERACION DE TAREAS
"""
dutiesDf = dutiesDf.dropna()
dutiesList = dutiesDf.values.tolist()

dutiesChanged = []
def searchForAreas(currentArea):
    if(isinstance(currentArea, float)): #convert in case is empty cell
        currentArea = ''
    assigned = currentArea.split('||')
    possibleError = False
    randomDuty = ''
    if (assigned[0] == ''):
        randomDuty = random.choice(
                    [duty for duty in dutiesList])
    else:
        if (len(dutiesList) == 1 and randomDuty in dutiesList):
            possibleError = True

        
        if(not possibleError):
            randomDuty = random.choice(
                        [duty for duty in dutiesList])
            while randomDuty in assigned:
                randomDuty = random.choice(
                        [duty for duty in dutiesList])

    if (possibleError):
        return False

    assigned.append(randomDuty)
    if (randomDuty in dutiesList):
        dutiesList.remove(randomDuty)

    return assigned


for idx in range(len(inputAsigned)):
    idxrow = inputAsigned.iloc[idx]
    assignedArray = []
    
    if (len(dutiesList) > 0):
        assignedArray = searchForAreas(idxrow['Tarea_asignada'])

    if (assignedArray):
        joinedArray = "||".join(assignedArray)
        inputAsigned.loc[idx, 'Tarea_asignada'] = joinedArray
        dutiesChanged.append(True)
    else:
        dutiesChanged.append(False)

inputAsigned['assigned_ocured_area'] = dutiesChanged
print('AREAS GENERADAS CORRECTO ')

"""
ARRAY CON HISTORIAL PARA GUARDAR EN EL EXCEL 
"""
historyDf = inputAsigned.copy()
historyDf.drop('assigned_ocured_person', axis=1, inplace=True)
historyDf.drop('assigned_ocured_area', axis=1, inplace=True)

print(historyDf)
historyDf.to_excel('historial_noTouch.xlsx', index=False)



"""
Filtrar para tener solo lo necesario 
"""

#Filtrar y guardar areas 
filteredAreas = inputAsigned.loc[inputAsigned['assigned_ocured_area'] == True].copy()
filteredAreas.drop('Persona_asignada', axis=1, inplace=True)
filteredAreas.drop('assigned_ocured_person', axis=1, inplace=True)

filteredAreas['tarea_asignada'] = filteredAreas['Tarea_asignada'].apply(lambda x: x.split('||')[-1] if '||' in x else x)
filteredAreas.drop('Tarea_asignada', axis=1, inplace=True)
filteredAreas.drop('assigned_ocured_area', axis=1, inplace=True)

filteredAreas.to_excel('filtrado_tareas.xlsx', index=False)


#Filtrar y guardar personas 
filteredPersons = inputAsigned.loc[inputAsigned['assigned_ocured_person'] == True].copy()
filteredPersons.drop('Tarea_asignada', axis=1, inplace=True)
filteredPersons.drop('assigned_ocured_area', axis=1, inplace=True)
print(filteredPersons)
filteredPersons['persona_asignada'] = filteredPersons['Persona_asignada'].apply(lambda x: x.split('||')[-1] if '||' in x else x)
filteredPersons.drop('Persona_asignada', axis=1, inplace=True)
filteredPersons.drop('assigned_ocured_person', axis=1, inplace=True)

filteredPersons.to_excel('filtrado_personas.xlsx', index=False)
