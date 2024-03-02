import os
import csv
import codecs
from time import sleep
from base64 import b64decode
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service



# --------------------------------------------------



sizes = {
    'default': {
        'w': 1050,
        'h': 718
    },
    'desktop': {
        'w': 1366,
        'h': 768
    },
    'desktop/2': {
        'w': 683,
        'h': 728
    },
    'mobile': {
        'w': 320,
        'h': 711
    },
    'mobile-landscape': {
        'w': 711,
        'h': 320
    },
    'square': {
        'w': 620,
        'h': 620
    }
}

service = Service('./chromedriver.exe')
options = Options()

optionSize = sizes['default']
# options.add_argument('--headless')
options.add_argument('--incognito')
options.add_argument(f'window-size={optionSize['w']},{optionSize['h']}')
options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(options, service)
driver.get('https://boletim-escolar.vercel.app')



# --------------------------------------------------



def AcceptCookieConcent():
    cookieName = '__cookie_concent'
    cookieConcent = driver.get_cookie(cookieName)

    def setCookieConcent():
        driver.add_cookie({'name': '__cookie_concent', 'value': 'true'})
        return AcceptCookieConcent()

    if cookieConcent == None: setCookieConcent()

    cookieConcent = bool(driver.get_cookie(cookieName).get('value'))

    if cookieConcent == True: return
    else: setCookieConcent()

def OpenCloseAside(intention: str):
    if intention not in ['open', 'close']:
        raise ValueError('\n\tIntention must be "open" or "close".')

    if intention == 'open':
        menuButton = driver.find_element(By.CSS_SELECTOR, '[aria-label="Abrir Menu"]')
    else:
        menuButton = driver.find_element(By.CSS_SELECTOR, '[aria-label="Fechar Menu"]')
    return menuButton.click()

def OpenCloseDetails(index: int):
    if index not in [0, 1, 2, 3, 4]:
        raise ValueError('\n\tIndex must be (0, 1, 2, 3 or 4).')

    asideContent = driver.find_element(By.ID, 'aside-content')
    detailsList = asideContent.find_elements(By.TAG_NAME, 'details')

    detailsList[index].click()

def ConfigureSchoolCard():
    OpenCloseAside('open')
    OpenCloseDetails(0)

    asideContent = driver.find_element(By.ID, 'aside-content')
    detailsList = asideContent.find_elements(By.TAG_NAME, 'details')
    details_EnableDisable = detailsList[0]
    bimonthlyButtonList = details_EnableDisable.find_elements(By.TAG_NAME, 'button')

    for bimonthlyButton in bimonthlyButtonList[1:4]:
        bimonthlyButton.click()
    bimonthlyButtonList[5].click()

    OpenCloseAside('close')

def CreateSchoolCard():
    schoolName = 'E. E. Emmanuel Oliveira Lopes'
    teacherName = 'Ana Luiza dos Santos'

    def Create(student: list, gradesAndAbsences: list, numberOfExecutions: int):
        if numberOfExecutions != 1:
            BypassingWebsiteBug()

        def InsertOrganizationInfo():
            inputSchool = driver.find_element(By.ID, 'school')
            inputTeacher = driver.find_element(By.ID, 'teacher')

            try:
                inputSchool.send_keys(schoolName)
                inputTeacher.send_keys(teacherName)
            except:
                raise Exception('\n\tUnable to enter value.')
        def InsertStudentInfo():
            inputStudentName         = driver.find_element(By.ID, 'student.name')
            inputStudentNumber       = driver.find_element(By.ID, 'student.number')
            inputStudentyearAndClass = driver.find_element(By.ID, 'student.yearAndClass')

            inputGradesFirstBimesterOfPortuguese        = driver.find_element(By.ID, 'studentAcademicRecord.Português.grades.firstQuarter')
            inputGradesFirstBimesterOfMathematics       = driver.find_element(By.ID, 'studentAcademicRecord.Matemática.grades.firstQuarter')
            inputGradesFirstBimesterOfArt               = driver.find_element(By.ID, 'studentAcademicRecord.Artes.grades.firstQuarter')
            inputGradesFirstBimesterOfSciences          = driver.find_element(By.ID, 'studentAcademicRecord.Ciências.grades.firstQuarter')
            inputGradesFirstBimesterOfHistória          = driver.find_element(By.ID, 'studentAcademicRecord.História.grades.firstQuarter')
            inputGradesFirstBimesterOfGeography         = driver.find_element(By.ID, 'studentAcademicRecord.Geografia.grades.firstQuarter')
            inputGradesFirstBimesterOfPhysicalEducation = driver.find_element(By.ID, 'studentAcademicRecord.Educação Física.grades.firstQuarter')

            inputAbsencesFirstBimesterOfPortuguese        = driver.find_element(By.ID, 'studentAcademicRecord.Português.absences.firstQuarter')
            inputAbsencesFirstBimesterOfMathematics       = driver.find_element(By.ID, 'studentAcademicRecord.Matemática.absences.firstQuarter')
            inputAbsencesFirstBimesterOfArt               = driver.find_element(By.ID, 'studentAcademicRecord.Artes.absences.firstQuarter')
            inputAbsencesFirstBimesterOfSciences          = driver.find_element(By.ID, 'studentAcademicRecord.Ciências.absences.firstQuarter')
            inputAbsencesFirstBimesterOfHistória          = driver.find_element(By.ID, 'studentAcademicRecord.História.absences.firstQuarter')
            inputAbsencesFirstBimesterOfGeography         = driver.find_element(By.ID, 'studentAcademicRecord.Geografia.absences.firstQuarter')
            inputAbsencesFirstBimesterOfPhysicalEducation = driver.find_element(By.ID, 'studentAcademicRecord.Educação Física.absences.firstQuarter')

            studentName, studentNumber, studentYear, studentClass = student
            studentYearAndClass = (studentYear + ' ' + studentClass)

            half_index = len(gradesAndAbsences) // 2
            grades = gradesAndAbsences[:half_index]
            absences = gradesAndAbsences[half_index:]

            try:
                inputStudentName.send_keys(studentName)
                inputStudentNumber.clear()
                inputStudentNumber.send_keys(studentNumber)
                inputStudentyearAndClass.send_keys(studentYearAndClass)

                inputGradesFirstBimesterOfPortuguese.clear()
                inputGradesFirstBimesterOfPortuguese.send_keys(grades[0])
                inputGradesFirstBimesterOfMathematics.clear()
                inputGradesFirstBimesterOfMathematics.send_keys(grades[1])
                inputGradesFirstBimesterOfArt.clear()
                inputGradesFirstBimesterOfArt.send_keys(grades[2])
                inputGradesFirstBimesterOfSciences.clear()
                inputGradesFirstBimesterOfSciences.send_keys(grades[3])
                inputGradesFirstBimesterOfHistória.clear()
                inputGradesFirstBimesterOfHistória.send_keys(grades[4])
                inputGradesFirstBimesterOfGeography.clear()
                inputGradesFirstBimesterOfGeography.send_keys(grades[5])
                inputGradesFirstBimesterOfPhysicalEducation.clear()
                inputGradesFirstBimesterOfPhysicalEducation.send_keys(grades[6])

                inputAbsencesFirstBimesterOfPortuguese.clear()
                inputAbsencesFirstBimesterOfPortuguese.send_keys(absences[0])
                inputAbsencesFirstBimesterOfMathematics.clear()
                inputAbsencesFirstBimesterOfMathematics.send_keys(absences[1])
                inputAbsencesFirstBimesterOfArt.clear()
                inputAbsencesFirstBimesterOfArt.send_keys(absences[2])
                inputAbsencesFirstBimesterOfSciences.clear()
                inputAbsencesFirstBimesterOfSciences.send_keys(absences[3])
                inputAbsencesFirstBimesterOfHistória.clear()
                inputAbsencesFirstBimesterOfHistória.send_keys(absences[4])
                inputAbsencesFirstBimesterOfGeography.clear()
                inputAbsencesFirstBimesterOfGeography.send_keys(absences[5])
                inputAbsencesFirstBimesterOfPhysicalEducation.clear()
                inputAbsencesFirstBimesterOfPhysicalEducation.send_keys(absences[6])
            except:
                raise Exception('\n\tUnable to enter value.')

        if numberOfExecutions == 1:
            InsertOrganizationInfo()
        InsertStudentInfo()

        GenerateImage()
        GetImage()

    csvStudents = 'students.csv'
    encoding = 'utf-8'
    with codecs.open(csvStudents, 'r', encoding) as csvFile:
        csvReader = csv.reader(csvFile)

        student = []
        numberOfStudent = 0
        gap = 4

        for index, row in enumerate(csvReader):
            if index != 0:
                student.append(row[0])

                if index == (gap-3): # Pega somente linha 1 de cada aluno
                    gradesAndAbsences = row[1:]

                if index == gap:
                    gap += 4
                    numberOfStudent += 1

                    Create(student, gradesAndAbsences, numberOfStudent)

                    # print(f'{numberOfStudent}° aluno:\n\t{student}\n\t{gradesAndAbsences}\n')
                    student.clear()

def GenerateImage():
    generateImageButton = driver.find_element(By.ID, 'generate-image')
    generateImageButton.click()
    sleep(3)
    script = 'const overlayCard = document.querySelector(".swal2-container");if(overlayCard){return overlayCard.remove();}'
    driver.execute_script(script)

def ConvertBase64ToImage(imageBase64: str):
    try:
        if not imageBase64.startswith('data:image/png;base64,'):
            raise ValueError('The base64 provided does not appear to be a PNG image.')

        imageBase64 = imageBase64[len('data:image/png;base64,'):]
        imageBytes = b64decode(imageBase64)

        return imageBytes
    except ValueError as ve: print(f'\n\t{ve}')
    except Exception as e: print(f'\n\tError converting base64 to image: {str(e)}')

def GetImage():
    OpenCloseAside('open')
    OpenCloseDetails(2)

    asideContent = driver.find_element(By.ID, 'aside-content')
    detailsList = asideContent.find_elements(By.TAG_NAME, 'details')
    details_Images = detailsList[2]

    downloadImageButton = details_Images.find_element(By.TAG_NAME, 'a')
    imageTitle = downloadImageButton.get_attribute('download')
    imageBase64 = downloadImageButton.get_attribute('href')

    imageBytes = ConvertBase64ToImage(imageBase64)

    os.makedirs('school_report_cards', exist_ok=True)
    with open(f'school_report_cards/{imageTitle}', 'wb') as imageFile:
        imageFile.write(imageBytes)
        print(f'Image {imageTitle} successfully saved.')

    ExcludeImageFromSite(details_Images)
    OpenCloseAside('close')

def ExcludeImageFromSite(details_Images):
    removeImageButton = details_Images.find_element(By.TAG_NAME, 'button')
    removeImageButton.click()

def BypassingWebsiteBug():
    sleep(0.15)
    OpenCloseAside('open')
    OpenCloseDetails(0)
    asideContent = driver.find_element(By.ID, 'aside-content')
    detailsList = asideContent.find_elements(By.TAG_NAME, 'details')
    details_EnableDisable = detailsList[0]
    bimonthlyButtonList = details_EnableDisable.find_elements(By.TAG_NAME, 'button')

    bimonthlyButtonList[0].click() # Double click
    bimonthlyButtonList[0].click()

    OpenCloseAside('close')



# --------------------------------------------------



sleep(2) # Tempo do Loader do site

AcceptCookieConcent()
ConfigureSchoolCard()
CreateSchoolCard()

sleep(2)
driver.quit()