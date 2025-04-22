import pdfplumber
import pandas as pd


class EInvoiceConverter:
    def __init__(self, pdf_paths):
        self.pdf_paths = pdf_paths
        self.columns =["Sira No", "Mal Hizmet", "Miktar", "Iskonto Orani", "Iskonto Tutari", "KDV Orani", "KDV Tutari", "Diğer Vergiler", "Hizmet Tutari"]
        self.alternative_columns  = [["Sira No", "Sira"], ["Mal Hizmet", "Hizmet Açıklamasi", "Açiklama" ], ["Miktar","Mktr"], ["Iskonto Orani", "Isknt %"], "Iskonto Tutari", ["KDV Orani", "KDV %"], "KDV Tutari", "Diğer Vergiler", ["Hizmet Tutari", "Net Tutar"]]
        self.data = {col: [] for col in self.columns}
        self.structured_table = []
        self.structured_nested_list = []

    def set_structure(self):
        nested_list = []
        for p in range(0, len(self.pdf_paths)):
            with pdfplumber.open(self.pdf_paths[p]) as pdf:            
                for page in pdf.pages:
                    text = page.extract_tables()
                
                    for table in text:
                        if table and table[0]:
                            table[0] = [item.replace("\n", " ") if item is not None else "" for item in table[0]]       
                        
                        if "Sira" in table[0][0] or "Sıra" in table[0][0]:
                            structured_table = []
                            for row in table:
                                structured_table.append(row)
                                
                            structured_table[0] = [item.replace("ı", "i").replace("İ", "I") for item in structured_table[0]]
                            
                            structured_temp_tables_list = structured_table              
                            self.structured_nested_list.append(structured_temp_tables_list)
                            break

        alternative_columns = self.alternative_columns
        structured_table = self.structured_nested_list
        for b in range(0,len(structured_table)):
            z = 0
            for col in alternative_columns:
                found = False
                if z == 0:
                    for i in range(0, len(alternative_columns[z])):
                        for j in range(0,len(structured_table[b][0])):
                            if alternative_columns[z][i] in structured_table[b][0][j]:
                                found = True
                                for a in range(1,len(structured_table[b])):
                                    self.data[col[0]].append(structured_table[b][a][j])
                                break
                        if found == True:
                            break

                elif z == 1:
                    for i in range(0, len(alternative_columns[z])):
                        for j in range(0,len(structured_table[b][0])):
                            if alternative_columns[z][i] in structured_table[b][0][j]:
                                found = True
                                for a in range(1,len(structured_table[b])):
                                    self.data[col[0]].append(structured_table[b][a][j])
                                break
                        if found == True:
                            break
                elif z == 2:
                    for i in range(0, len(alternative_columns[z])):
                        for j in range(0,len(structured_table[b][0])):
                            if alternative_columns[z][i] in structured_table[b][0][j]:
                                found = True
                                for a in range(1,len(structured_table[b])):
                                    self.data[col[0]].append(structured_table[b][a][j])
                                break
                        if found == True:
                            break

                elif z == 3:
                    for i in range(0, len(alternative_columns[z])):
                        for j in range(0,len(structured_table[b][0])):
                            if alternative_columns[z][i] in structured_table[b][0][j]:
                                found = True
                                for a in range(1,len(structured_table[b])):
                                    self.data[col[0]].append(structured_table[b][a][j])
                                break
                        if found == True:
                            break    

                elif z == 5:
                    for i in range(0, len(alternative_columns[z])):
                        for j in range(0,len(structured_table[b][0])):
                            if alternative_columns[z][i] in structured_table[b][0][j]:
                                found = True
                                for a in range(1,len(structured_table[b])):
                                    self.data[col[0]].append(structured_table[b][a][j])
                                break
                        if found == True:
                            break    

                elif z == 8:
                    for i in range(0, len(alternative_columns[z])):
                        for j in range(0,len(structured_table[b][0])):
                            if alternative_columns[z][i] in structured_table[b][0][j]:
                                found = True
                                for a in range(1,len(structured_table[b])):
                                    self.data[col[0]].append(structured_table[b][a][j])
                                break
                        if found == True:
                            break      
                else:
                    for i in range(0, len(structured_table[b][0])):
                        if col in structured_table[b][0][i]:
                            found = True
                            for j in range(1,len(structured_table[b])):
                                self.data[col].append(structured_table[b][j][i])
                            break
                if found == False:
                    for _ in range(1, len(structured_table[b])):
                        if z == 1 or z == 2 or z == 3 or z ==5 or z == 8:
                            self.data[col[0]].append(" ")
                        else:
                            self.data[col].append(" ")    
                z+=1
        df = pd.DataFrame(self.data)

        df = df.replace("\n", " ", regex=True)
        df = df[df["Sira No"].notna() & (df["Sira No"] != "")]
        df = df[df["Mal Hizmet"].notna() & (df["Mal Hizmet"] != "")]
        df = df[df['Hizmet Tutari'].notna() & (df['Hizmet Tutari'].astype(str).str.strip() != '')]

        return df
    
    def create_excel(self, dataframe, excel_name):
        dataframe.to_excel(f'{excel_name}.xlsx', index=False, engine='openpyxl')    
