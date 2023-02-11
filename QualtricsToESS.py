import wx
import os
import pandas as pd
import json
import platform
from collections import OrderedDict

# row in df containing json values (e.g. {"ImportId": QID1_1}) - note Pandas take first row as headers
json_row_num = 1

def ballot_list_to_excel(ballots, file_path, file_name):
    """Takes in the ballots, which are a list of lists, and writes to an excel sheet in election ES&S CVR format"""
        
    df = pd.DataFrame(ballots)
    # make column names Choice 1, Choice 2, etc.
    col_names = []
    for col_idx in range(1, len(df.columns) + 1):
        col_names.append('Choice ' + str(col_idx))
    df.columns = col_names

    df.insert(0, 'Cast Vote Record', range(1, len(ballots) + 1))
    df.insert(1, 'Precinct', file_name)
    df.insert(2, 'Ballot Style', 'Qualtrics')

    writer = pd.ExcelWriter(file_path)
    df.to_excel(writer, sheet_name='ballots', index=False, header=True)
    writer.save()

def convert_to_ballots(df, progress_dialog):
    """Convert the dataframe to ballots (a list of lists)"""
    # the first row index containing ballots
    ballot_idx_start = 2
    num_ballots_total = len((df[ballot_idx_start:]).index)

    ballots = []*len(df)

    # loop through each row (ballot)
    for row_idx, row_contents in df[ballot_idx_start:].iterrows():
        ballot = [] # list of candidates
        result_dict = OrderedDict() # key: cell contents (ranking), value: candidate given that ranking

        # loop through columns within the row, put into result_dict if not empty
        for col_idx in range(len(df.columns)):
            if (row_contents[col_idx] != ''):
                result_dict[row_contents[col_idx]] = (str(df.iloc[0, col_idx]))

        # sort result_dict by the key (ranking)
        # e.g. ballot = {1: Option 3, 2: Option 4, 3: Option 2, 4: Option 1}
        ballot = OrderedDict(sorted(result_dict.items()))

        # take the values (the candidates)
        ballot = list(ballot.values())
        ballots.append(ballot)

        ballot_count = row_idx - ballot_idx_start + 1
        # want to reset progress_dialog for each election
        # so subtract 1 in order to not reach max, else ProgressDialog closes
        progress_dialog.Update(int((ballot_count / num_ballots_total) * 100) - 1)
    return(ballots)

def make_json_config(xls_filepath, contest_name, candidates, output_dir):
    """Creates JSON configuration file and writes it to output_dir"""
    candidate_map_list = []*len(candidates)
    for candidate in candidates:
        candidate_map_list.append({
            "name": candidate,
            "code": "",
            "excluded": False
        })

    output_dict = {
        "tabulatorVersion": "1.3.0",
        "outputSettings": {
            "contestName": contest_name,
            "outputDirectory": "..\\RCTab_Output\\" + 'RCTab_Output_' + contest_name,
            "tabulateByPrecinct": False,
            "generateCdfJson": False
        },
        "cvrFileSources": [{
            "provider": "ess",
            "filePath": xls_filepath,
            "contestId": "",
            "firstVoteColumnIndex": "4",
            "firstVoteRowIndex": "2",
            "idColumnIndex": "1",
            "precinctColumnIndex": "2",
            "overvoteDelimiter": "",
            "overvoteLabel": "overvote",
            "undervoteLabel": "undervote",
            "undeclaredWriteInLabel": "",
            "treatBlankAsUndeclaredWriteIn": False
        }],
        "candidates": candidate_map_list,
        "rules": {
            "tiebreakMode": "previousRoundCountsThenRandom",
            "overvoteRule": "exhaustImmediately",
            "winnerElectionMode": "singleWinnerMajority",
            "numberOfWinners": "1",
            "decimalPlacesForVoteArithmetic": "4",
            "maxSkippedRanksAllowed": "unlimited",
            "maxRankingsAllowed": "max",
            "randomSeed": "1234"
        }
    }
    output_json_path = output_dir + '\\' + contest_name + '.json'
    with open(output_json_path, 'w', encoding='utf-8') as file:
        json.dump(output_dict, file, indent=4)
    return(output_json_path)

def qualtrics_to_ess(input_csv, progress_dialog, output_dir, rctab_dir):
    '''For each election in the input_csv, exports a ESS/CVR format Excel file and the json_config file into output_dir'''
    df = pd.read_csv(input_csv)
    # ignore row if all null
    df.dropna(how='all',inplace=True)
    # replace na with empty string
    df = df.fillna('')

    # get indices of question columns (column name starts with Q and contains underscore)
    qcol_idx_list = []
    # check the third row containing json values 
    json_row = df.iloc[json_row_num].values

    for idx in range(len(json_row)):        
        import_id_value = json.loads(json_row[idx])["ImportId"]
        if import_id_value.startswith("Q") and "_" in import_id_value and "_TEXT" not in import_id_value:
            qcol_idx_list.append(idx)
            # rename json cell to keep only question ID (Q1D1_1 to become Q1D1)
            df.iloc[json_row_num, idx] = import_id_value.split("_")[0]

    # keep only the question columns
    df = df.iloc[:,qcol_idx_list]       

    # row containing the candidates
    candidate_row_num = 0
    # exported Qualtrics CSV is in the format "[question text] - [candidate name]"
    # keep only the candidate name
    for idx in range(len(df.columns)):
        df.iloc[candidate_row_num, idx] = " - ".join(df.iloc[candidate_row_num, idx].split(" - ")[1:])
    

    # sorted set of unique column names
    unique_json_col_names = sorted(set(df.iloc[json_row_num]))
    for col_name in unique_json_col_names:
        # columns for this question, e.g. Q1_1, Q1_2, Q1_3
        q_col_names = df.iloc[json_row_num][df.iloc[json_row_num] == col_name].index
        # get only columns with this value
        sub_df = pd.DataFrame(df.loc[:,q_col_names])
        # get what's to the left of underscore in the first question cell of row
        election_name = q_col_names[0].split('_')[0]
        
        filename = input_csv.split('\\')[-1].replace('.csv', '')
        xls_filepath = output_dir + '\\' + filename + '_' + election_name + '.xlsx'
        
        candidates = list(sub_df.iloc[0])
        contest_name = filename + '_' + election_name
        output_json_path = make_json_config(xls_filepath, contest_name, candidates, output_dir)
        progress_dialog.Update(0, "Reading ballots for election: " + election_name)
        ballot_list_to_excel(convert_to_ballots(sub_df, progress_dialog), xls_filepath, filename + '_' + election_name)

        # run RCTab via command line
        rctab_dir += "\\rcv\\bin"
        os.chdir(rctab_dir)
        os.system("rcv -cli " + '"' + output_json_path + '"')
    return(output_dir)

class WindowNew(wx.Dialog):
    def __init__(self, *args, **kwds):
        wx.Dialog.__init__(self, *args, **kwds)

        self.label_candidate_file = None
        self.button_create = None
        self.max_progress_dialog_value = 100

        self.show_ui()
    
    def show_ui(self):
        self.label_candidate_file = wx.StaticText(self, wx.ID_ANY, "Qualtrics CSV File")
        self.text_ctrl_candidate_file = wx.TextCtrl(self, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_candidate_file_browse = wx.Button(self, wx.ID_ANY, "Browse")
        self.Bind(wx.EVT_BUTTON, self.ui_browse_candidate_file, self.button_candidate_file_browse)

        self.button_create = wx.Button(self, wx.ID_ANY, "Convert")
        self.Bind(wx.EVT_BUTTON, self.ui_convert, self.button_create)
        self.button_create.Enable(False)

        self.sizer_main = wx.FlexGridSizer(4, 1, 5, 0)

        self.row_0 = wx.FlexGridSizer(1, 3, 5, 5)
        self.row_0.Add(self.label_candidate_file, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT | wx.LEFT, 5)
        self.row_0.Add(self.text_ctrl_candidate_file, 0, wx.EXPAND, 0)
        self.row_0.Add(self.button_candidate_file_browse, 0, wx.RIGHT, 5)
        self.row_0.AddGrowableCol(1)
        self.sizer_main.Add(self.row_0, 1, wx.EXPAND, 0)

        self.row_1 = wx.FlexGridSizer(1, 3, 10, 10)
        self.row_1.Add(wx.StaticText(self, wx.ID_ANY, "Output Directory"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT | wx.LEFT, 5)
        self.m_dirPickerOutputDir = wx.DirPickerCtrl( self, wx.ID_ANY, "", u"Select the output directory", wx.DefaultPosition, wx.Size( 400,-1 ), wx.DIRP_DEFAULT_STYLE )
        self.m_dirPickerOutputDir.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
        self.row_1.Add(self.m_dirPickerOutputDir, 0, wx.RIGHT, 5)
        self.sizer_main.Add(self.row_1)

        self.row_2 = wx.FlexGridSizer(1, 3, 10, 10)
        self.row_2.Add(wx.StaticText(self, wx.ID_ANY, "RCTab Directory "), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT | wx.LEFT, 5)
        self.m_dirPickerRCTab = wx.DirPickerCtrl( self, wx.ID_ANY, os.getcwd() + '\\rctab_v1.3.0_windows', u"Select 'rctab_v1.3.0_windows' directory containing RCTab", wx.DefaultPosition, wx.Size( 400,-1 ), wx.DIRP_DEFAULT_STYLE )
        self.m_dirPickerRCTab.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
        self.row_2.Add(self.m_dirPickerRCTab, 0, wx.RIGHT, 0)
        self.sizer_main.Add(self.row_2)

        self.sizer_main.Add(self.button_create, 0, wx.ALIGN_RIGHT | wx.BOTTOM | wx.RIGHT, 5)

        self.SetSizer(self.sizer_main)
        self.sizer_main.Fit(self)
        self.sizer_main.AddGrowableRow(0)
        self.sizer_main.AddGrowableCol(0)
        self.Layout()

        self.SetTitle("Qualtrics CSV to RCTab")

    def ui_browse_candidate_file(self, event):
        election_candidate_file = wx.FileDialog(self, "", os.getcwd(), "", "CSV file (*.csv)|*.csv|All files|*.*", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if election_candidate_file.ShowModal() == wx.ID_CANCEL:
            return
        self.set_candidate_file(election_candidate_file.GetPath())
        self.ui_check_complete()

    def set_candidate_file(self, candidate_file):
        self.candidate_file = candidate_file
        self.text_ctrl_candidate_file.SetValue(self.candidate_file)

    def ui_convert(self, event):
        if not self.is_valid_csv(self.candidate_file):
            wx.MessageDialog(self, "Unable to load Qualtrics CSV file. Please verify that the file specified is the correct file.", caption="Load Error", style=wx.OK | wx.ICON_ERROR | wx.CENTRE).ShowModal()
        elif not os.path.exists(self.m_dirPickerOutputDir.GetPath()):
            wx.MessageDialog(self, "Output directory path does not exist.", caption="Error", style=wx.OK | wx.ICON_ERROR | wx.CENTRE).ShowModal()
        elif not os.path.exists(self.m_dirPickerRCTab.GetPath() + '\\rcv\\bin\\rcv'):
            wx.MessageDialog(None, "Invalid path for RCTab Directory. Please select the folder 'rctab_v1.3.0_windows' for the RCTab Directory.", caption="Conversion failure", style=wx.OK | wx.ICON_ERROR | wx.CENTRE).ShowModal()
        else:
            try:
                progress_dialog = wx.ProgressDialog("Processing Ballots", "", maximum=self.max_progress_dialog_value, parent=self, style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE | wx.PD_ELAPSED_TIME | wx.PD_ESTIMATED_TIME | wx.PD_REMAINING_TIME)
                progress_dialog.Fit()
                
                # intermediate_dir contains the Excel file and config JSON file for each election, used as imput to RCTab
                intermediate_dir = self.m_dirPickerOutputDir.GetPath() + '\\Intermediate_Output'
                if not os.path.exists(intermediate_dir):
                    os.makedirs(intermediate_dir)
                
                qualtrics_to_ess(self.candidate_file, progress_dialog, intermediate_dir, self.m_dirPickerRCTab.GetPath())
                
                progress_dialog.Update(self.max_progress_dialog_value)
                progress_dialog.Destroy()
                
                rctab_output_dir = self.m_dirPickerOutputDir.GetPath() + '\\RCTab_Output'
                if wx.MessageDialog(self,
                "Open tabulated output directory? \n\nOutputted to: \n\n" + rctab_output_dir,
                caption="Conversion complete", style=wx.YES_NO | wx.YES_DEFAULT | wx.ICON_QUESTION).ShowModal() == wx.ID_YES:
                    
                    if platform.system() == "Windows":
                        # os.startfile(self.m_dirPickerOutputDir.GetPath())
                        os.startfile(self.m_dirPickerOutputDir.GetPath() + '\\RCTab_Output')
        
                self.Close()
            except Exception as e:
                progress_dialog.Destroy()
                print(e)
                wx.MessageDialog(self, "An unexpected error occurred.", caption="Error", style=wx.OK | wx.ICON_ERROR | wx.CENTRE).ShowModal()

    def is_valid_csv(self, input_csv):
        '''
        Validates the format of the input CSV candidate_file by checking if the 
        json row contains a key of "ImportId" with a value that starts with Q
        and contains an underscore
        '''
        df = pd.read_csv(input_csv)

        # check the row containing json values 
        json_row = df.iloc[json_row_num].values
        for cell in json_row:
            try:
                import_id_value = json.loads(cell)['ImportId']
            except Exception as e:
                pass
            else:
                if import_id_value.startswith("Q") and "_" in import_id_value and "_TEXT" not in import_id_value:
                    return True
        
        return False

    def ui_check_complete(self):
        self.button_create.Enable(not not (self.candidate_file))

def main():
    app = wx.App()
    app_new_ui = WindowNew(None)
    app_new_ui.ShowModal()
    app_new_ui.Destroy()
    app.MainLoop()

if __name__ == "__main__":
    main()
