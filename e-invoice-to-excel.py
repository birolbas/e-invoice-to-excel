import pdfplumber
import pandas as pd
pdf_path = ["C:/Users/bbasb/Downloads/yamanticaret.pdf","C:/Users/bbasb/Downloads/fatura.pdf","C:/Users/bbasb/Downloads/keskinoğlu.pdf","C:/Users/bbasb/Downloads/kasap.pdf","C:/Users/bbasb/Downloads/dasdas.pdf","C:/Users/bbasb/Downloads/voest.pdf","C:/Users/bbasb/Downloads/asd.pdf"]
all_columns = []
columns  = ["Sira No", "Mal Hizmet", "Miktar", "Iskonto Orani", "Iskonto Tutari", "KDV Orani", "KDV Tutari", "Diğer Vergiler", "Hizmet Tutari"]
alternative_columns  = ["Sira No", ["Mal Hizmet", "Hizmet Açıklamasi", "Açiklama" ], "Miktar", "Iskonto Orani", "Iskonto Tutari", "KDV Orani", "KDV Tutari", "Diğer Vergiler", ["Hizmet Tutari", "Net Tutar"]]
structured_temp_tables_list = []
structured_nested_list = []
structured_table = []
data = {col: [] for col in columns}

def extract_tables():
    for p in range(0, len(pdf_path)):
        with pdfplumber.open(pdf_path[p]) as pdf:            
           for page in pdf.pages:
                text = page.extract_tables()
               
                for table in text:
                    if table and table[0]:
                        table[0] = [item.replace("\n", " ") if item is not None else "" for item in table[0]]       
                    
                    if "Sira No" in table[0][0] or "Sıra No" in table[0][0]:
                        structured_table = []
                        for row in table:
                            structured_table.append(row)
                            
                        structured_table[0] = [item.replace("ı", "i").replace("İ", "I") for item in structured_table[0]]
                        
                        structured_temp_tables_list = structured_table              
                        structured_nested_list.append(structured_temp_tables_list)
                        break
                
    return structured_nested_list
          
def set_structure(nested_list):
    structured_table = nested_list
    for b in range(0,len(nested_list)):
        z = 0
        for col in alternative_columns:
            found = False
            if z == 1:
                for i in range(0, len(alternative_columns[z])):
                    for j in range(0,len(structured_table[b][0])):
                        if alternative_columns[z][i] in structured_table[b][0][j]:
                            found = True
                            for a in range(1,len(structured_table[b])):
                                data[col[0]].append(structured_table[b][a][j])
                            break
                    if found == True:
                        break
            elif z == 8:
                for i in range(0, len(alternative_columns[z])):
                    for j in range(0,len(structured_table[b][0])):
                        if alternative_columns[z][i] in structured_table[b][0][j]:
                            found = True
                            for a in range(1,len(structured_table[b])):
                                data[col[0]].append(structured_table[b][a][j])
                            break
                    if found == True:
                        break      
            else:
                for i in range(0, len(structured_table[b][0])):
                    if col in structured_table[b][0][i]:
                        found = True
                        for j in range(1,len(structured_table[b])):
                            data[col].append(structured_table[b][j][i])
                        break
            if found == False:
                for _ in range(1, len(structured_table[b])):
                    if z == 1 or z == 8:
                        data[col[0]].append(" ")
                    else:
                        data[col].append(" ")    
            z+=1
    return data
                
def create_dataframe(data):
    df = pd.DataFrame(data)
    df = df.replace("\n", " ", regex=True)
    df = df[df["Sira No"].notna() & (df["Sira No"] != "")]
    df = df[df["Mal Hizmet"].notna() & (df["Mal Hizmet"] != "")]
    return df

def create_excel(dataframe):
    df.to_excel('veriler.xlsx', index=False, engine='openpyxl')                
    return "DONE"
extracted_nested_list = extract_tables()
data_for_dataframe = set_structure(nested_list=extracted_nested_list )
df = create_dataframe(data=data_for_dataframe)
answer = create_excel(dataframe = df)
