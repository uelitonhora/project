import PySimpleGUI as sg

class MainApp:
    def __init__(self):

        # Theme
        sg.theme('darkblue17')

        # Layouts
        layout = [
            [sg.Text('A.Telecom', font='Arial 18 bold',
                     text_color='white', size=(605, 1), justification='center')],

            [sg.Frame(layout=[
            [sg.Radio('Gerar Planilhas', "RADIO1",key='csv',pad=(20,0),default=True, enable_events=True), sg.Radio('Gerar txt', "RADIO1",key='txt', pad=(20,2),enable_events=True)]], size=(390,1),  pad=(160,0), title='Selecione:')],
            [sg.T(size=(20, 1))],
            [sg.Text('Selecione a planilha:',key='txt_name', font='Arial 10')],
            [sg.Input(size=(70, 2), key='txt_dir', background_color='white'),
            sg.FileBrowse('Pesquisar', key='dir_files', file_types=(("Planilha Excel", "*.xlsx"),), size=(50, 1)),
            sg.FolderBrowse('Pesquisar', key= 'dest_folder', size=(50, 1), target='txt_dir')],
            [sg.Text('Diretório de Destino:', font='Arial 10')],
            [sg.Input(size=(70, 2), key='dir_folder', background_color='white'), sg.FolderBrowse(
                'Pesquisar', size=(50, 1))],
            [sg.T(size=(20, 2))],
            [sg.Button('Processar Arquivos', font='Arial 10 bold', size=(25, 2)), sg.T(size=(20, 2)),
             sg.Button('Resetar', font='Arial 10 bold', size=(25, 2))],
            [sg.Text("0%", size=(625, 1), font='Arial 10',
                     key='percent', justification='center')],
            [sg.ProgressBar(1, orientation='h', size=(
                605, 30),  key='progress', pad=(5, 5), bar_color=('green', 'gray'))]
        ]

        # Create the Window
        self.window = sg.Window('A.Telecom', layout,
                                size=(625, 330), icon=r'logo.ico')

    def Start(self):
        while True:
            event, values = self.window.read()

            if event == sg.WINDOW_CLOSED:
                break
            # Reset
            if event == 'Resetar':
                self.window['txt_dir'].update('')
                self.window['dir_folder'].update('')
                continue

            # csv create selected
            if event == 'csv':
                self.window['dest_folder'].update(visible=False)
                self.window['dir_files'].update(visible=True)
                self.window['txt_dir'].update('')
                self.window['txt_name'].update('Selecione a Planilha:')

            # txt create selected
            if event == 'txt':
                self.window['dest_folder'].update(visible=True)
                self.window['dir_files'].update(visible=False)
                self.window['txt_dir'].update('')
                self.window['txt_name'].update('Diretório com as planilhas:')
                
                
            if event == 'Processar Arquivos':
                if values['csv']:
                    # Verification
                    if values['txt_dir'] == '':
                        sg.popup(
                            'Planilha não selecionada.', title='Atenção!')
                        continue

                    # Verification
                    elif values['dir_folder'] == '':
                        sg.popup(
                            'O diretório destino não foi informado.', title='Atenção!')
                        continue

                    file, folder_dest = values['txt_dir'], values['dir_folder']
                    utils.create_csv(file, folder_dest)
                    sg.popup(
                            'Relatórios gerado com sucesso!', title='Sucesso!')  
                    self.window['txt_dir'].update('')
                    self.window['dir_folder'].update('')                 

                else:
                    # Verification
                    if values['txt_dir'] == '':
                        sg.popup(
                            'Diretório com as planilhas não informado.', title='Atenção!')
                        continue

                    # Verification
                    elif values['dir_folder'] == '':
                        sg.popup(
                            'O diretório destino não foi informado.', title='Atenção!')
                        continue
                    
                    folder_files, folder_dest = values['txt_dir'], values['dir_folder']
                    files_names = utils.find_csv(folder=folder_files)
                    if len(files_names) == 0:
                        sg.popup(
                            'No diretório informado não há Planilhas.', title='Atenção!')
                        continue
                    utils.create_txt(files_names, folder_dest)
                    sg.popup(
                            'Relatórios gerado com sucesso!', title='Sucesso!')
                    self.window['txt_dir'].update('')
                    self.window['dir_folder'].update('')


tela = MainApp()
tela.Start()
