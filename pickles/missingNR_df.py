from openpyxl.utils import absolute_coordinate
import pandas as pd

def coord(list_of_dicts):
    list_of_names = ["General Information", "Environmental Disclosures", "Social Disclosures", "Governance Disclosures", "Fuel Converter"]
    for dict in list_of_dicts:
        name = list_of_names[list_of_dicts.index(dict)]
        for key in dict.keys():
            for i in range(len(dict[key])):
                dict[key][i] = absolute_coordinate(dict[key][i])
                dict[key][i] = name + "!" + dict[key][i]
    return list_of_dicts

dict_missingNR_GI = {
    "1.0.0": ["D6:E8", "D10:E12", "E99", "H107", "C139", "E159"],
    "1.0.1": ["D6:E8", "D10:E12", "E99", "H107", "C139", "E159"],
    "1.1.0": ["D6:E8", "D10:E12", "E99", "H108", "C140", "E160"],
    "1.1.1": ["D6:E8", "D10:E12", "E440", "H449", "C556", "E576"]
}

dict_missing_ED = {
    "1.0.0": ["I59:K61","D30:F44", "G30:I44", "J30:J44", "D57", "G10", "J19", "G29", "G74", "G124", "D175", "G214", "E235"],
    "1.0.1": ["I59:K61","D30:F44", "G30:I44", "J30:J44", "D57", "G10", "J19", "G29", "G74", "G124", "D175", "G214", "E235"],
    "1.1.0": ["I59:K61","D30:F44", "G30:I44", "J30:J44", "D57", "G10", "J19", "G29", "G74", "G124", "D175", "G214", "E235"],
    "1.1.1": ["I59:K61","D30:F44", "G30:I44", "J30:J44", "D57", "G10", "J19", "G29", "G74", "G184", "D309", "G433", "E539"]
}

dict_missingNR_SD = {
    "1.0.0": ["D43:H45", "D50", "D57:H58", "D60", "D74:H75"],
    "1.0.1": ["D43:H45", "D50", "D57:H58", "D60", "D74:H75"],
    "1.1.0": ["D44:H46", "D51", "D58:H59", "D61", "D75:H76"],
    "1.1.1": ["D134:H136", "D141", "D148:H149", "D151", "D165:H166"]
}

dict_missingNR_GD = {
    "1.0.0": ["H5", "H10", "H22:H26", "H30:H32"],
    "1.0.1": ["H5", "H10", "H22:H26", "H30:H32"],
    "1.1.0": ["H5", "H10", "H22:H26", "H30:H32"],
    "1.1.1": ["H5", "H10", "H22:H26", "H30:H32"]
}

dict_missingNR_Fuel = {
    "1.0.0": ["B10", "L10", "P10", "T10", "X10", "AB10", "AF10", "AJ10", "AN10", "AR10", "B16", "L16", "P16", "T16", "X16", "AB16", "AF16", "AJ16", "AN16", "AR16", "B17", "L17", "P17", "T17", "X17", "AB17", "AF17", "AJ17", "AN17", "AR17"],
    "1.0.1": ["B10", "L10", "P10", "T10", "X10", "AB10", "AF10", "AJ10", "AN10", "AR10", "B16", "L16", "P16", "T16", "X16", "AB16", "AF16", "AJ16", "AN16", "AR16", "B17", "L17", "P17", "T17", "X17", "AB17", "AF17", "AJ17", "AN17", "AR17"],
    "1.1.0": ["A10", "A11", "A12", "A13", "A14", "A15", "A16", "A17", "A18", "A19", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19"],
    "1.1.1": ["A10", "A11", "A12", "A13", "A14", "A15", "A16", "A17", "A18", "A19", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19"]
}

updated = coord([dict_missingNR_GI, dict_missing_ED, dict_missingNR_SD, dict_missingNR_GD, dict_missingNR_Fuel])
df_missingNR_GI = pd.DataFrame(updated[0])
df_missing_ED = pd.DataFrame(updated[1])
df_missingNR_SD = pd.DataFrame(updated[2])
df_missingNR_GD = pd.DataFrame(updated[3])
df_missingNR_Fuel = pd.DataFrame(updated[4])

df_missingNR_SD.loc[5] = {
    "1.0.0": "Social Disclosures!None",
    "1.0.1": "Social Disclosures!None",
    "1.1.0": "Social Disclosures!$E$27",
    "1.1.1": "Social Disclosures!$E$27"
}

missingNR_df = pd.concat([df_missingNR_GI, df_missing_ED, df_missingNR_SD, df_missingNR_GD, df_missingNR_Fuel], ignore_index=True)
# missingNR_df.to_pickle("pickles/missingNR_df.pkl")