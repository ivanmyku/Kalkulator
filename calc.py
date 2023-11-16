# install OpenPyXl,streamlit
import streamlit as st
from openpyxl import Workbook, load_workbook
import inquirer
import datetime


def chose_cell(filepath, quarter, cell, user_input):
    wb = load_workbook(filepath)
    ws = wb[quarter]
    try:
        if ws[cell].value is not None:
            ws[cell] = float(ws[cell].value) + float(user_input)
        else:
            ws[cell].value = 0
            wb.save(filepath)
    except ValueError:
        st.warning('Please enter and numeric value')
    wb.save(filepath)
    st.info(f'Stan wydatków na jedzenie w {now.strftime("%B")} wynosi {ws[cell].value}', icon="ℹ️")

def submit():
    st.session_state.my_text = st.session_state.receipt
    st.session_state.receipt = ""

# load in your workbook
FilePath = r'D:\Finanse\Finansy zycia.xlsx'
now = datetime.date.today()
st.title("Kalkulator wydatków na jedzenie")

inp = st.text_input(label="Input", placeholder="Dodaj nowy paragon...",
                      key='receipt')

if inp is not '':
    st.error("Czy na pewno chcesz wprowadzić paragon?")
    if st.button("Nie"):
        exit()
    elif st.button("Tak"):
        match now.strftime('%B'):
            case "November":
                chose_cell(FilePath, 'Wydatki', 'B1', inp)
            case "December":
                chose_cell(FilePath, 'Wydatki', 'B2', inp)
            case "January":
                chose_cell(FilePath, 'Wydatki', 'B3', inp)
            case "February":
                chose_cell(FilePath, 'Wydatki', 'B4', inp)
            case "March":
                chose_cell(FilePath, 'Wydatki', 'B5', inp)
            case "April":
                chose_cell(FilePath, 'Wydatki', 'B6', inp)
            case "May":
                chose_cell(FilePath, 'Wydatki', 'B7', inp)
            case "June":
                chose_cell(FilePath, 'Wydatki', 'B8', inp)
            case "July":
                chose_cell(FilePath, 'Wydatki', 'B9', inp)
            case "August":
                chose_cell(FilePath, 'Wydatki', 'B10', inp)
            case "September":
                chose_cell(FilePath, 'Wydatki', 'B11', inp)
            case "October":
                chose_cell(FilePath, 'Wydatki', 'B12', inp)
else:
    exit()

