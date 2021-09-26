import os

import matplotlib.pyplot as plt
import pandas as pd

from fpdf import FPDF

BASE_DIR = 'student_data/'
STUDENT_PICTURES = 'Pics for assignment/'

LIGHT_RED = '#FF7F7F'
LIGHT_GREEN = '#90EE90'


def generate_score_sheet(reg_no, df):
    # filter out data corresponding to current registration id
    score_sheet = df.loc[df['Reg. No.'] == reg_no]
    score_sheet = score_sheet.drop(['Reg. No.'], axis=1)

    # build score table
    cell_text = []
    cell_colors = []
    for row in range(len(score_sheet)):
        cell_text.append(score_sheet.iloc[row])
        if score_sheet.iloc[row]['Outcome'] == 'Correct':
            cell_colors.append(['w'] + ([LIGHT_GREEN] * 5))
        elif score_sheet.iloc[row]['Outcome'] == 'Incorrect':
            cell_colors.append(['w'] + ([LIGHT_RED] * 5))
        else:
            cell_colors.append(['w'] * 6)
    # plt.figure()
    table = plt.table(cellText=cell_text, colLabels=score_sheet.columns, cellColours=cell_colors, loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(16)
    table.scale(2.5, 2.5)
    plt.axis('off')

    # save score tables of each student as an image
    image_folder = BASE_DIR + '{}/images/'.format(reg_no)
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)
    plt.savefig(image_folder + '{}.jpg'.format('score_sheet'), dpi=300, bbox_inches='tight')


def generate_report_card(reg_no, df):
    data = df.loc[df['Registration Number'] == reg_no]
    # print(data)
    student_folder = BASE_DIR + '{}/'.format(reg_no)
    if not os.path.exists(student_folder):
        os.makedirs(student_folder)

    full_name = data['First Name'].iloc[0] + " " + data['Last Name'].iloc[0]
    school = data['Name of School'].iloc[0]
    country = data['Country of Residence'].iloc[0]
    date_of_test = data['Date and time of test'].iloc[0]
    total_marks_scored = data['Your score'].sum(skipna=True)
    total_max_marks = data['Score if correct'].sum(skipna=True)
    final_result = data['Final result'].iloc[0]
    round_no = data['Round'].iloc[0]
    percentage = round((total_marks_scored / total_max_marks) * 100, 2)

    student_image_path = 'Pics for assignment/{}.jpg'.format(full_name)
    student_score_sheet_path = BASE_DIR + '{}/images/score_sheet.jpg'.format(reg_no)

    pdf = FPDF('P', 'mm', 'A4')
    pdf.add_page()

    pdf.set_font('Arial', '', 24)
    pdf.image('logo.jpg', w=30, type='jpg')

    pdf.set_xy(x=pdf.get_x() + 60, y=10)

    pdf.cell(90, 15, 'Institute Name', 0, 1, 'C')

    pdf.set_xy(x=pdf.get_x() + 60, y=pdf.get_y() - 3)

    pdf.set_font('Arial', '', 14)
    pdf.cell(90, 15, 'Institute Address', 0, 1, 'C')

    pdf.set_xy(x=pdf.get_x() + 60, y=pdf.get_y() - 7)

    pdf.set_font('Arial', '', 10)
    pdf.cell(90, 15, 'Institute Contact Number', 0, 1, 'C')

    pdf.set_y(50)

    pdf.line(10, 45, 200, 45)

    pdf.set_font('Arial', '', 14)
    pdf.cell(200, 5, 'REPORT CARD', 0, 1, 'C')
    pdf.ln(5)
    pdf.image(student_image_path, w=30, h=30, type='jpg')

    pdf.set_xy(x=pdf.get_x() + 40, y=pdf.get_y() - 25)

    pdf.set_font('Arial', 'B', 10)
    pdf.multi_cell(w=150, h=5, txt="REG. NO.:\nNAME:\nSCHOOL:\nCOUNTRY:\nDATE OF TEST:", align='L')

    pdf.set_xy(x=pdf.get_x() + 70, y=pdf.get_y() - 25)

    pdf.set_font('Arial', '', 10)
    pdf.multi_cell(w=150, h=5, txt="{}\n{}\n{}\n{}\n{}".format(reg_no, full_name, school, country, date_of_test), align='L')

    pdf.set_xy(x=pdf.get_x() + 120, y=pdf.get_y() - 24)

    pdf.set_font('Arial', 'B', 10)
    pdf.multi_cell(w=150, h=5, txt="ROUND:\nMARKS SCORED:\nMAX MARKS:\nPERCENTAGE:", align='L')

    pdf.set_xy(x=pdf.get_x() + 160, y=pdf.get_y() - 20)

    pdf.set_font('Arial', '', 10)
    pdf.multi_cell(w=150, h=5, txt="{}\n{}\n{}\n{}".format(round_no, total_marks_scored, total_max_marks, percentage), align='L')

    pdf.set_xy(x=pdf.get_x() + 15, y=pdf.get_y() + 10)

    pdf.set_font('Arial', 'B', 13)
    pdf.cell(165, 5, 'RESULT: ', 0, 1, 'C')
    pdf.ln(7)
    pdf.set_font('Arial', '', 13)

    pdf.set_x(pdf.get_x() + 15)

    pdf.cell(160, 5, '{}'.format(final_result), 0, 1, 'C')
    pdf.image(student_score_sheet_path, x=15, y=125, w=180, type='jpg')

    pdf.output(student_folder + 'report.pdf', 'F')


if __name__ == '__main__':

    # PREPROCESS DATA

    file = "Dummy Data.xlsx"
    df = pd.read_excel(file, header=1, index_col=None)

    # select required columns for generating score tables
    df_for_score_sheet = df[['Registration Number', 'Question No.', 'What you marked', 'Correct Answer',
                             'Outcome (Correct/Incorrect/Not Attempted)', 'Score if correct', 'Your score']]

    df_for_report = df[['Round', 'First Name', 'Last Name', 'Registration Number', 'Grade', 'Name of School', 'Gender',
                        'Date of Birth', 'City of Residence', 'Date and time of test', 'Country of Residence',
                        'Score if correct', 'Your score', 'Final result']]

    # rename columns for convenience
    df_for_score_sheet = df_for_score_sheet.rename(
        columns={'Registration Number': 'Reg. No.', 'Outcome (Correct/Incorrect/Not Attempted)': 'Outcome',
                 'What you marked': 'Your answer'})

    # fill null values with a '-'
    df_for_score_sheet['Your answer'].fillna('-', inplace=True)

    # GENERATE SCORE TABLES

    # loop for each registration id
    for i in pd.unique(df_for_score_sheet['Reg. No.']):
        generate_score_sheet(i, df_for_score_sheet)

        # create report cards
        generate_report_card(i, df_for_report)
