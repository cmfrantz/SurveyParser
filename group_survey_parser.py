# -*- coding: utf-8 -*-
"""
Created on Sun Dec 13 10:58:25 2020

@author: cariefrantz
"""



####################
# IMPORTS
####################
import os
from tkinter import Tk, filedialog
from statistics import mean, stdev
import numpy as np
import pandas as pd

####################
# VARIABLES
####################

# General
NA_LIST = [np.nan, 'NA', 'na', 'N/A', 'n/a', 'N/a', 'nan', '']

# Rows containing actual students in the Canvas gradebook csv
GB_HEAD = 0
GB_FIRSTROW = 1
GB_LASTROW = -1
GB_KEEP_COLS = ['Student', 'ID', 'SIS User ID', 'SIS Login ID',
                   'Root Account', 'Section']
GB_SID = 'SIS Login ID'

# Survey Response sheet map
survey_sheets = {
    'responses'     : {
        'sheet'     : 'Form Responses 1',
        'header'    : 0
        },
    'map_questions' : {
        'sheet'     : 'ResponseMap',
        'header'    : 3
        },
    'map_points'    : {
        'sheet'     : 'PointMap',
        'header'    : 3
        }
    }

####################
# FUNCTIONS
####################

##########
# MISCELLANEOUS
##########

def split_email(email):
    '''
    Splits student ID off of student email address
    
    Parameters
    ----------
    email : str
        Student university email address
        
    Returns
    -------
    student_id : str
        Student university ID
    '''
    student_id = email.split('@')[0].lower()
    return student_id

def add_empty_cols(dataframe, cols):
    # Adds empty columns with designated headers to the dataframe
    for col in cols:
        dataframe[col]=''
    return dataframe


def add_prefix_suffix(col_list, add_str, p_or_s):
    '''
    Adds prefix or suffix to newhead columns for the gradebook

    Parameters
    ----------
    col_list : list of str
        List of newhead columns.
   add_str : str
        String to add as prefix or suffix.
    p_or_s : str
        'p' = prefix
        's' = suffix

    Returns
    -------
    col_list : list of str
        List of columns with prefix or suffix added.

    '''
    if p_or_s == 'p':
        col_list = [add_str + ': ' + col for col in col_list]
    elif p_or_s == 's':
        col_list = [col + ' (' + add_str + ')' for col in col_list]
    return col_list


def unique_peers(survey_map):
    '''
    Finds the list of peers (e.g., 'Peer1', 'Peer2') used in the survey map

    Parameters
    ----------
    survey_map : pandas.DataFrame
        Dataframe mapping survey questions to categories.

    Returns
    -------
    peer_list : list of str
        List of unique peers from the survey map.

    '''
    peer_list = list(set(list(survey_map[
        survey_map['student'].str.contains(r'peer')].loc[:,'student'])))
    peer_list.sort()
    return peer_list


def first_peer(survey_map):
    firstpeer = unique_peers(survey_map)[0]
    return firstpeer


def find_columns(survey_map, student, category='any', map_col = 'newhead',
                prefix=False, suffix=False):
    '''
    Finds and returns names of columns based on input parameters

    Parameters
    ----------
    survey_map : pandas.DataFrame
        DataFrame that maps survey questions to categories.
    student : str
        Value from the student column of the survey_map.
        Possible values include 'general', 'self', and 'peer1'
    category : str, optional
        Value from the category column of the survey_map.
        The default is 'any', which returns all columns for the student.
    map_col : str, optional
        The desired column name type to return. The default is 'newhead'.
        The other option is 'survey_column'.
    prefix : str, optional
        Use this option to add a prefix to the column headers,
        e.g., 'SE' becomes 'SE: Column Name'. The default is False.
    suffix: str, optional
        Use this option to add a suffix to the column headers,
        e.g., 'avg' becomes 'Column Name (avg)'. The default is False.

    Returns
    -------
    cols : list of str
        List of column names.

    '''
    if category != 'any':
        cols = list(survey_map[
            (survey_map.student.str.contains(student))
            & (survey_map.category.str.contains(category))][map_col])
    else:
        cols = list(survey_map[
            (survey_map.student.str.contains(student))][map_col])
    cols = [col for col in cols if str(col) != 'nan']
    if prefix:
        cols = add_prefix_suffix(cols, prefix, 'p')
    elif suffix:
        cols = add_prefix_suffix(cols, suffix, 's')
    return cols


def gen_pe_rating_columns(rating_cols):
    cols=[]
    for col in rating_cols:
        cols.append(col + ' (avg)')
        cols.append(col + ' (std)')
    return cols

##########
# FILE IMPORT & PREPROCESSING
##########
def import_sheet(text='Select file', directory=os.getcwd(), filetype='csv',
                header=0, xls_sheet_map=[]):
    '''
    Imports spreadsheet file as a pandas DataFrame

    Parameters
    ----------
    text : str, optional
        Text to display in the file selection dialog.
        The default is 'Select file'.
    directory : str, optional
        The directory to search. The default is os.getcwd().
    filetype : str, optional
        The type of file. The default is 'csv'. Other options are 'xls'.
    header : int, optional
        The row location of the header in the file to import. The default is 0.
    xls_sheet_map : dict, optional
        Dictionary (of dictionaries) including all of the sheets
        (for an Excel file) to load. The dictionary should contain the sheet
        names to load as dictionary names, with sub-dictionaries containing
        'name' (the name of the variable to map to the sheet) and
        'header' (the header row for the sheet). The default is [].

    Returns
    -------
    data : pandas.DataFrame or dict of pandas.DataFrame
        DataFrame containing the loaded spreadsheet data, or
        dictionary of DataFrames containing data for each of the specified
        Excel sheets.
    filepath : str
        Filepath for the opened file

    '''
    file_types = {
        'csv'   :   [('CSV', '*.csv')],
        'xls'   :   [('Excel', '.xls .xlsx')]
        }
    # Open a file selection dialog
    root=Tk()
    fpath = filedialog.askopenfilename(
        title=text, initialdir=directory, filetypes=file_types[filetype])
    root.destroy()
    # Load and return the file
    if filetype=='csv':
        data = pd.read_csv(fpath, sep=',', header=header)
    elif filetype=='xls':
        if not xls_sheet_map:
            data = pd.read_excel(filepath, header=header)
        else:
            data = {}
            for sheet in xls_sheet_map:
                sheet_info = xls_sheet_map[sheet]
                data[sheet] = pd.read_excel(
                    filepath, sheet_name=sheet_info['sheet'],
                    header = sheet_info['header'])

    return data, fpath


def prep_gradebook(grade_book):
    # Trim gradebook list of student names and Canvas info
    grade_book = (
        grade_book.iloc[GB_FIRSTROW:GB_LASTROW][GB_KEEP_COLS])

    # Add an unformatted name column
    namelist = []
    for name in grade_book['Student']:
        name_split = name.split(', ')
        namelist.append(' '.join((name_split[1], name_split[0])))
    grade_book['Name'] = namelist

    return grade_book


def prep_map(survey_results, survey_map):
    # Remap the map with the imported column names
    survey_map['survey_column'] = survey_results.columns
    return survey_map


##########
# PROCESS EVALUATION SURVEYS
##########

def process_self_evals(grade_book, survey_results, survey_map, map_points):
    '''
    Find self-evaluations in the survey results and parse them

    Parameters
    ----------
    grade_book : pandas.DataFrame
        Pre-processed gradebook.
    survey_results : pandas.DataFrame
        Imported results from the Google Survey.
    survey_map : pandas.DataFrame
        Imported results from the Google Survey.
    map_points : dict
        Dictionary mapping survey responses to points.

    Returns
    -------
    grade_book : pandas.DataFrame
        Gradebook with the self-evaluation information added.

    '''

    # Make list of self-evaluation questions (include general info)
    cols_self = (find_columns(survey_map, 'self', map_col='survey_column')
                 + find_columns(survey_map, 'general', map_col='survey_column'))
    evals_self = survey_results.loc[:,cols_self].copy()
    evals_self.columns = (find_columns(survey_map, 'self', prefix='SE')
                          + find_columns(survey_map, 'general'))

    # Add the self-evaluation columns to the gradebook
    grade_book = add_empty_cols(grade_book, evals_self.columns)

    # Convert ratings to points
    rating_cols = find_columns(survey_map, 'self', 'rating', prefix='SE')
    evals_self.loc[:,rating_cols] = convert_ratings(
        map_points, evals_self[rating_cols])

    # Identify student based on university email address (login ID)
    # and fill in self-evaluation data
    # Find the survey column that contains the reporting student's email address
    [col_email] = find_columns(survey_map,'self','email', prefix = 'SE')
    list_emails = evals_self[col_email]
    for email in list_emails:
        # Split out the student ID and match in the gradebook
        student_id = split_email(email)
        if student_id in list(grade_book[GB_SID]):
            grade_book.loc[
                grade_book[GB_SID]==student_id,
                evals_self.columns] = (
                    evals_self.loc[evals_self[col_email]==email].values)
        else:
            [student_name] = evals_self.loc[evals_self[col_email]==email,'SE: Name']
            student_name = fix_name(student_name)
            if student_name not in list(grade_book['Name']):
                eval_row = survey_results.loc[survey_results[find_columns(
                    survey_map, 'self', 'email',
                    map_col='survey_column')[0]]==email].copy()
                student_name = find_student(grade_book, survey_map, eval_row, 'self')
            if student_name not in NA_LIST:
                grade_book.loc[
                    grade_book['Name']==student_name,
                evals_self.columns] = (
                    evals_self.loc[evals_self[col_email]==email].values)
    # Check gradebook for missing values and enter nan
    check_cols = rating_cols + find_columns(survey_map, 'self', 'score', prefix='SE')
    grade_book.loc[:,check_cols]=grade_book[check_cols].replace('', np.nan)

    return grade_book


def process_peer_evals(grade_book, survey_results, survey_map, map_points):
    '''
    Find peer-evaluations in the survey results and parse them

    Parameters
    ----------
    gradebook : pandas.DataFrame
        Pre-processed gradebook.
    survey_results : pandas.DataFrame
        Imported results from the Google Survey.
    survey_map : pandas.DataFrame
        Imported results from the Google Survey.
    map_points : dict
        Dictionary mapping survey responses to points.

    Returns
    -------
    gradebook : pandas.DataFrame
        Gradebook with the peer-evaluation information added.

    '''

    # Make a list of all peer-evaluations
    cols_peer_all = survey_map[survey_map['student'].str.contains(r'peer')]
    all_peers = list(set(list(cols_peer_all['student'])))
    firstpeer = first_peer(survey_map)
    newcols_peers = list(
        survey_map.loc[survey_map['student']==firstpeer,'newhead'])

    evals_peer = pd.DataFrame(columns=['review_row'] + newcols_peers)
    for peer in all_peers:
        peer_evals = survey_results.loc[
            :,survey_map.loc[survey_map['student']==peer, 'survey_column']]
        peer_evals = np.column_stack((peer_evals.index,peer_evals.values))
        evals_peer = evals_peer.append(pd.DataFrame(
            data=peer_evals, columns = evals_peer.columns), ignore_index=True)

    # Add the peer evaluation columns to the gradebook
    pe_cols = {
        'general'   : ['PE: N'],
        'comments'  : find_columns(
            survey_map, first_peer(survey_map), 'comments'),
        'rating'    : find_columns(
            survey_map, first_peer(survey_map), 'rating'),
        'all'       : []
        }
    pe_cols['rating_avg'] = gen_pe_rating_columns(pe_cols['rating'])
    for col_set in [col_set for col_set in pe_cols if col_set not in ('rating', 'all')]:
        pe_cols['all'] = pe_cols['all'] + pe_cols[col_set]

    grade_book = add_empty_cols(grade_book, pe_cols['all'])

    # Set up the evaluation collection per student
    evals_peer_compiled = {}
    for student in list(grade_book['Name']):
        evals_peer_compiled[student] = pd.DataFrame(columns=evals_peer.columns).copy()
        evals_peer_compiled[student].drop_duplicates(inplace=True)

    # For the peer evalulations, try to identify student based on the entered
    #   first and last name
    [col_peername] = find_columns(survey_map, first_peer(survey_map), 'name')
    for peer_eval in evals_peer.index:
        # Try to find the name in the gradebook
        name = evals_peer.loc[peer_eval, col_peername]
        if name in NA_LIST:
            continue
        if name not in list(grade_book['Name']):
            name = fix_name(name)
            if name not in list(grade_book['Name']):
                eval_row = survey_results.loc[
                    evals_peer.loc[peer_eval,'review_row']].copy()
                name = find_student(grade_book, survey_map, eval_row, name)
        if name == 'NA':
            continue
        # Add the evaluation to the compiled set
        evals_peer_compiled[name] = evals_peer_compiled[name].append(
            evals_peer.loc[peer_eval])

    # Calculate averages, compile comments, and add to gradebook
    for student in evals_peer_compiled:
        eval_df = evals_peer_compiled[student].copy()
        if not eval_df.empty:
            avgs=average_ratings(evals_peer_compiled[student].copy(), pe_cols,
                               map_points)
            grade_book.loc[grade_book['Name']==student, pe_cols['all']]=avgs.values

    # Check gradebook for missing values and enter nan
    grade_book.loc[:,pe_cols['all']]=grade_book[pe_cols['all']].replace('', np.nan)

    return grade_book


def average_ratings(eval_df, pe_cols, map_points):
    '''
    Compile multiple ratings into averages, stdevs, and combined comments

    Parameters
    ----------
    eval_df : pandas.DataFrame
        DESCRIPTION.
    pe_cols : list of str
        List of all peer evaluation columns.
    map_points : dict
        Dictionary mapping rating responses to points.

    Returns
    -------
    avgs : pandas.DataFrame
        Dataframe containing compiled responses.

    '''
    # Create holder DataFrame
    avgs = pd.DataFrame(columns=pe_cols['all'])
    avgs.loc[0,'PE: N'] = eval_df.shape[0]

    # Combine comments
    for comment_col in pe_cols['comments']:
        comments = list(eval_df[comment_col])
        comments = [c for c in list(eval_df[comment_col]) if c not in NA_LIST]
        avgs[comment_col] = ' | '.join(comments)

    # Convert the ratings
    ratings = convert_ratings(map_points, eval_df[pe_cols['rating']])

    # Average the ratings
    for c in pe_cols['rating']:
        avgs[c + ' (avg)'] = round(mean(list(ratings[c])),2)
        if avgs.loc[0,'PE: N']==1:
            avgs[c + ' (std)'] = np.nan
        else:
            avgs[c + ' (std)'] = round(stdev(list(ratings[c])),2)

    return avgs


def convert_ratings(map_points, ratings):
    '''
    Converts categorical ratings to points

    Parameters
    ----------
    map_points : pandas.DataFrame
        Dataframe mapping survey questions to catebories.
    ratings : pandas.DataFrame
        Dataframe containing the columns containing ratings.

    Returns
    -------
    ratings : pandas.DataFrame
        Dataframe containing point values for the given ratings.

    '''
    ratings=ratings.copy()
    for i in ratings.index:
        for c in ratings.columns:
            ratings.loc[i,c] = map_points[ratings.loc[i,c]]
    return ratings


def calc_differences(grade_book, survey_map):
    # List columns
    se_cols = find_columns(survey_map, 'self', 'rating')
    pe_cols = find_columns(survey_map, first_peer(survey_map), 'rating')
    match_cols = set(se_cols).intersection(pe_cols)
    for col in match_cols:
        diff_col_name = 'SE-PE: ' + col
        grade_book[diff_col_name] = (
            grade_book['SE: '+col] - grade_book[col + ' (avg)'])
    return grade_book

def fix_name(name):
    # Get rid of trailing/leading spaces
    name = name.strip()
    # Initial caps
    name = ' '.join([name.capitalize() for name in name.split(' ')])
    return name

def find_student(grade_book, survey_map, eval_row, student):
    '''
    When student cannot be automatically matched in the gradebook,
    prompt the user to enter the student's name based on other identifying
    information.'

    Parameters
    ----------
    gradebook : pandas.DataFrame
        Gradebook dataframe.
    survey_map : pandas.DataFrame
        DESCRIPTION.
    eval_row : Series
        Extracted row from the evaluation filled out by the reviewing student.
    student : str
        'self' or a peer's name.

    Returns
    -------
    name : str
        Student name (First Last) or 'NA' if not found.

    '''
    # Internal variables
    input_str_start = 'Who is this student? '
    input_str_end = ('\nEnter their name as listed in Canvas (First Last). '
                     + 'To skip this student or ignore this result,'
                     + ' enter "NA". > ')
    error_str = 'Student not found in gradebook: '

    # Find columns containing potentially identifying info
    r_info = {
        'section'   : '',
        'team'      : '',
        'email'     : '',
        'name'      : ''
        }
    for category in r_info:
        r_info[category] = eval_row[
            find_columns(survey_map, 'self', category=category,
                        map_col='survey_column')[0]]
    peer_names = []
    for peer in unique_peers(survey_map):
        peer_name = eval_row[
            find_columns(survey_map, peer, category='name',
                        map_col='survey_column')[0]]
        if not peer_name in NA_LIST:
            peer_names.append(peer_name)

    # Loop for self-evaluation
    if student == 'self':
        reviewer_id = split_email(r_info['email'])
        while True:
            name = input(
                input_str_start
                 + '"' + reviewer_id + '", "' + r_info['name']
                 + '" in Section ' + r_info['section']
                 + ', Team ' + r_info['team']
                 + input_str_end)
            if (name in list(grade_book['Name'])) or (name in NA_LIST):
                break
            print(error_str + name)

    # Loop for peer-evaluation
    else:
        while True:
            name = input(
                input_str_start
                + '"' + student + '"'
                 + '" in Section ' + r_info['section']
                 + ', Team ' + r_info['team']
                + '\nOther team members evaluated by this person: '
                + r_info['name'] + ', ' + ', '.join(peer_names)
                + input_str_end)
            if (name in list(grade_book['Name'])) or (name in NA_LIST):
                break
            print(error_str + name)

    return name


####################
# MAIN
####################
if __name__ == '__main__':

    # Import gradebook and survey response csv files
    gradebook, filepath = import_sheet(
        text = 'Select Canvas gradebook file', header = GB_HEAD)
    dirPath = os.path.dirname(filepath)
    survey_data, filepath = import_sheet(
        text='Select downloaded survey response file', directory=dirPath,
        filetype='xls', xls_sheet_map=survey_sheets)

    # Prep the gradebook
    gradebook = prep_gradebook(gradebook)

    # Prep the survey & points maps
    survey_data['map_questions'] = prep_map(
        survey_data['responses'], survey_data['map_questions'])
    survey_pointmap = dict(zip(survey_data['map_points']['Rating'],
                        survey_data['map_points']['Points']))

    # Process self evaluations
    gradebook = process_self_evals(
        gradebook, survey_data['responses'], survey_data['map_questions'],
        survey_pointmap)

    # Process peer evaluations
    # This still needs to get the gradebook headers renamed to 'PE: '
    gradebook = process_peer_evals(
        gradebook, survey_data['responses'], survey_data['map_questions'],
        survey_pointmap)

    # Calculate differences between self-evals and peer evals
    gradebook = calc_differences(gradebook, survey_data['map_questions'])

    # Save as a new Excel sheet
    filename = input(
        'Done processing. What would you like to name the processed file?  > ')
    gradebook.to_excel(dirPath + '/' + filename + '.xls')
    print('Saved! All done :)')

    # Future work:
    # Sort students by team first
