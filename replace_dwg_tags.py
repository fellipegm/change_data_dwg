# -*- coding: utf-8 -*-
"""
Modifica os dados de arquivos .dwg baseado em um arquivo de entrada
tags_to_change.xlsx

april 2020
@author: Fellipe Garcia Marques | fellipegm@gmail.com
"""
import os
import glob
import re
import sys
import pandas as pd
import ezdxf
import shutil




def main():
    change_data()


def change_data():
    # =============================================================================
    # # Converte os arquivos dwg para o formato dxf, caso necessário
    # =============================================================================
    
    current_dir = os.getcwd()
    directory = glob.glob('*DL*/')
    if len(directory) > 1:
        print("Apenas uma pasta de documento por vez, remova a pasta excedente")
        input("Pressione qualquer tecla para sair...")
        sys.exit(1)
    if len(directory) == 0:
        print("Não foram encontrados documentos para conversão de dados\n")
        input("Pressione Enter para sair...")
        sys.exit(1)
    
    if (os.path.isdir(os.path.join(current_dir, directory[0][0:-1] + "\DXF"))):
        shutil.rmtree(os.path.join(current_dir, directory[0][0:-1] + "\DXF"))
    if (os.path.isdir(os.path.join(current_dir, directory[0][0:-1]))):
        os.mkdir(os.path.join(current_dir, directory[0][0:-1] + "\DXF"))
    
    print("Convertendo o documento {0} para DXF\n".format(directory[0][0:-1]))
    config = '\"ACAD2010\" \"DXF\" \"0\" \"1\"'
    doc = "\"" + os.path.normpath(current_dir) + '\\' + os.path.normpath(directory[0][0:-1]) + "\""
    dest = doc[0:-1] + "\DXF\""
    try:
        os.chdir("C:\Program Files\ODA\ODAFileConverter_title 21.2.0")
        os.system("ODAFileConverter.exe " + doc + ' ' + dest + ' ' + config)
        os.chdir(current_dir)
    except OSError:
        print("O conversor de DWG para DXF não foi encontrado no diretório:\n" \
              "C:\Program Files\ODA\ODAFileConverter_title 21.2.0\n")
        sys.exit(0)
    
    
    # =============================================================================
    # # Importa os dados
    # =============================================================================
    def get_system(text="", data=""):
        if type(text) is not str:
            return ""
        match = re.match("(^[0-9]{4})([A-Z]{2,5})([0-9]{3,5}[A-Z]{0,2}\.{0,1}[0-9]{0,2}).*", text)
        if match is not None:
            if data == "system" and len(match.groups()) >= 3:
                return match[1]
            elif data == "type" and len(match.groups()) >= 3:
                return match[2]
            elif data == "loop" and len(match.groups()) >= 3:
                return match[3]
        return ""
        
    change_df = pd.read_excel(os.path.join(current_dir, "tags_to_change.xlsx"), "data_replace")
    change_df["sistema_original"] = change_df["original"].apply(lambda string : get_system(string, data="system"))
    change_df["tipo_original"] = change_df["original"].apply(lambda string : get_system(string, data="type"))
    change_df["malha_original"] = change_df["original"].apply(lambda string : get_system(string, data="loop"))
    change_df["sistema_destino"] = change_df["destino"].apply(lambda string : get_system(string, data="system"))
    change_df["tipo_destino"] = change_df["destino"].apply(lambda string : get_system(string, data="type"))
    change_df["malha_destino"] = change_df["destino"].apply(lambda string : get_system(string, data="loop"))
    
    new_doc = change_df[change_df["original"]=="NUMERO"]["destino"].iloc[0]
    save_path = os.path.join(current_dir, new_doc)
    os.mkdir(save_path)
    os.mkdir(os.path.join(save_path, "DXF"))
    # =============================================================================
    # # Abre arquivos e modifica
    # =============================================================================
    dxfs = glob.glob(os.path.normpath(current_dir) + '\\' + directory[0] + "DXF\*.dxf")
    print("Modificando os dados... aguarde...")
    for dxf in dxfs:
        dxf_reader = ezdxf.readfile(dxf)
        msp = dxf_reader.modelspace()
        
    # =============================================================================
    #     Modifica o carimbo
    # =============================================================================
        carimbo = msp.query("INSERT[name==\"CARIMBO A3\"]")[0]
        dests_carimbo = change_df[change_df["tipo"] == "carimbo"].copy()
        for attrib in carimbo.attribs:
            for index, tag in enumerate(dests_carimbo["original"]):
                if tag == attrib.dxf.tag:
                    attrib.dxf.text = dests_carimbo["destino"].iloc[index]
                    if tag == "TITULO":
                        dests_carimbo["original"].iloc[index] = "old"
                    break
        
    # =============================================================================
    #   Pega o número da página e filtra apenas os dados a serem trocados nesta página
    # =============================================================================
        for attrib in carimbo.attribs:
            if attrib.dxf.tag == "FOLHA":
                folha = int(attrib.dxf.text)
                break
            
    
        change_page = change_df[((change_df["pagina"] == folha) | (change_df["pagina"] == "all"))]
    # =============================================================================
    #   Altera os dados dos textos
    # =============================================================================
        texts = msp.query("TEXT MTEXT")
        dests_text = change_page[change_page["tipo"] == "texto"]
        for text in texts:
            for index, dest_txt in enumerate(dests_text["original"]):
                try:
                    if text.dxf.text == dest_txt:
                        text.dxf.text = dests_text["destino"].iloc[index]
                        change_df.loc[dests_text.index[index], "modificado"] = 'Sim'
                        break
                except:
                    pass
    
    
    # =============================================================================
    #   Altera os dados dos inserts    
    # =============================================================================
        inserts = msp.query("INSERT")
        dests_insert = change_page[change_page["tipo"] == "insert"]
        for insert in inserts:
            try:
                sequencial_aux = ""
                tag_aux = ""
                tipo_idx = 0
                malha_idx = 0
                if insert.attribs[0].dxf.tag == "TAG_INSTRUMENTO":
                    tag_aux = insert.attribs[0].dxf.text
                    sequencial_aux = insert.attribs[1].dxf.text
                    tipo_idx = 0
                    malha_idx = 1
                elif insert.attribs[0].dxf.tag == "SEQUENCIAL_INSTRUMENTO":
                    tag_aux = insert.attribs[1].dxf.text
                    sequencial_aux = insert.attribs[0].dxf.text
                    tipo_idx = 1
                    malha_idx = 0
            except IndexError:
                continue
            for index, dest_insert in enumerate(dests_insert["original"]):
                if tag_aux == dests_insert["tipo_original"].iloc[index] and \
                    sequencial_aux == dests_insert["malha_original"].iloc[index]:
                        insert.attribs[tipo_idx].dxf.text = dests_insert["tipo_destino"].iloc[index]
                        insert.attribs[malha_idx].dxf.text = dests_insert["malha_destino"].iloc[index]
                        change_df.loc[dests_insert.index[index], "modificado"] = 'Sim'
                        break
    
    
    # =============================================================================
    #   Salva o documento     
    # =============================================================================
        save_filename = re.sub("R11.*04[0-9]{2}", new_doc, dxf).replace(os.path.normpath(current_dir) + '\\', "")
        dxf_reader.saveas(os.path.join(current_dir, save_path, 'DXF\\', save_filename))
    
            
    # =============================================================================
    #   Escreve arquivo de logs  
    # =============================================================================
    change_df.to_excel("change_log.xlsx")
    
    # =============================================================================
    # Transforma de volta para .dwg
    # =============================================================================
    os.chdir("C:\Program Files\ODA\ODAFileConverter_title 21.2.0")
    config = '\"ACAD2010\" \"DWG\" \"0\" \"1\"'
    os.system("ODAFileConverter.exe " + \
              os.path.normpath(os.path.join(save_path, 'DXF')) + ' ' + \
              os.path.normpath(os.path.join(save_path)) + ' ' + config)
    os.chdir(current_dir)


if __name__ == '__main__':
    main()