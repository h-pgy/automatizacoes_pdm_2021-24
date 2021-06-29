import os
import pandas as pd
from pathlib import Path

PATH_ONE_DRIVE = Path(r'C:\Users\h-pgy\one_drive_prefs\OneDrive - Default Directory\Shared Documents\Estruturação do PDM 2021-2024\Elaboração PDM Versão Final')

class MapFileRenamer:

    def __init__(self, dir_origi, new_dir):

        self.dir_origi = self.get_one_drive_path(dir_origi)
        self.new_dir = self.solve_new_path(new_dir)
        self.control = self.read_controle()

    def solve_new_path(self, new_path):

        new_path_one_drive = self.get_one_drive_path(new_path)
        if not os.path.exists(new_path_one_drive):
            os.mkdir(new_path_one_drive)

        return new_path_one_drive

    def get_one_drive_path(self, path_or_file):

        return PATH_ONE_DRIVE /path_or_file

    def read_controle(self):

        path_controle_one_drive = self.get_one_drive_path(
            'Controle das Devolutivas.xlsx')

        return pd.read_excel(path_controle_one_drive)

    def list_arquivos(self, path_dir=None):

        if path_dir is None:
            path_dir = self.dir_origi

        return os.listdir(path_dir)

    def get_num_meta_file(self, f):

        return float(f.split('_')[0])

    def get_novo_num_meta(self, file, control = None):

        if control is None:
            control = self.control

        num_origi = self.get_num_meta_file(file)

        row = control[control['Meta' ]==num_origi]

        num = row['NOVO Nº Meta'].values[0]

        if pd.isnull(num):
            return 'sem_numero_novo' + str(num_origi)
        else:
            return int(num)

    def rename_file(self, file, novo_num, path_salvar = None):

        if path_salvar is None:
            path_salvar = self.new_dir

        end = file.split('_')[-1]
        f_name = str(novo_num) + '_' + end

        return os.path.join(path_salvar, f_name)

    def copy_bin_file(self, old_name, new_name):

        with open(old_name, 'rb') as f:
            data = f.read()
            with open(new_name, 'wb') as new_f:
                new_f.write(data)

    def __call__(self):


        for old_file in self.list_arquivos():
            old_path = os.path.join(self.dir_origi, old_file)
            novo_num = self.get_novo_num_meta(old_file)
            new_name = self.rename_file(old_file, novo_num)

            self.copy_bin_file(old_path, new_name)
            print(f'{new_name} criado')

if __name__ == "__main__":

    dir_origi = 'Mapas - Regionalização/V5 - final (numeração antiga)'
    new_dir = 'Mapas - Regionalização/V5 - final (numeração nova)'
    rename = MapFileRenamer(dir_origi, new_dir)
    rename()