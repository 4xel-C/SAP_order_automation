import pandas as pd
import sys

# Create constants to work with the dataframe consistently
CODE = "code"
MANUFACTURER = "manufacturer"
DESCRIPTION = "description"
CATEGORY = "category"

# Create keywords used in the Items classification for the application
SOLVENTS = [
    "ACET",
    "CHLORO",
    "ETH",
    "DMSO",
    "DIMETHYL",
    "HEPT",
    "TETRA",
    "PROP",
    "TOLUE",
]

CONSUMABLES = [
    "AIGU",
    "GANT",
    "PIPETT",
    "PASTE",
    "PARAF",
    "FLACON",
    "POUB",
    "ALU",
    "SOPA",
    "KIM",
    "RMN",
    "ESSAIS",
]

PURIFICATION = ["COLON"]

MISC = ["SABLE", "SILICE", "GRANU", "SODIUM", "JAVEL"]


class Item:
    """
    Class to handle all items in the excel file. recieve 1 argument:
    Path (string) to the corresponding Excel file.
    """


    def __init__(self, path):
        self.path = path

        # Create the DataFrame, clean and categorize items. using the class methods
        self.df = self.__load_df(path)
        self.df = self.__clean_df(self.df)
        self.df = self.__categorize_items(self.df)

        # Create categories dataframe
        self.solvents = self.df.loc[self.df[CATEGORY] == "solvent",[CODE, DESCRIPTION]]
        self.consumables = self.df.loc[self.df[CATEGORY] == "consumable",[CODE, DESCRIPTION]]
        self.purification = self.df.loc[self.df[CATEGORY] == "purification",[CODE, DESCRIPTION]]
        self.misc = self.df.loc[self.df[CATEGORY] == "miscelanous",[CODE, DESCRIPTION]]
        self.others = self.df.loc[self.df[CATEGORY] == "other",[CODE, DESCRIPTION]]

    @classmethod
    def __load_df(cls, path: str) -> pd.DataFrame:
        """
        Load the dataframes containing all the informations for the items.
        Take the path of the excel file as an input and ouput the dataframe.
        input the user to enter the path if data not correctly loaded by the application.
        """
        while True:
            try:
                df = pd.read_excel(path)
                break
            except FileNotFoundError:
                path = input("File not found, enter file path manually:")
        return df

    @classmethod
    def __clean_df(cls, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean the DataFrame by removing NaN values from manufacturer column, homogenizing values, and renaming columns
        used by the application to ensure consistency.
        """
        try:
            df.rename(
                columns={
                    df.columns[0]: CODE,
                    df.columns[1]: MANUFACTURER,
                    df.columns[5]: DESCRIPTION,
                }
            )
        except IndexError:
            print("Wrong dataframe format")
            sys.exit(1)

        # drop rows with empty "fabricant" column =>  No avaibility, change the NaN value for nmr tubes (so they  are not deleted)
        df.loc[df[DESCRIPTION].str.contains("RMN"), MANUFACTURER] = "No data"
        df = df.dropna(subset=MANUFACTURER)

        # Homogenize data strings in description column
        df[DESCRIPTION] = df[DESCRIPTION].str.replace("\n", " ")
        df[DESCRIPTION] = df[DESCRIPTION].str.capitalize()
        df[DESCRIPTION] = df[DESCRIPTION].str.strip()

        return df

    @classmethod
    def __categorize_items(cls, df: pd.DataFrame) -> pd.DataFrame:
        """
        Input a dataframe containing all items and create a columns "category" to categorize all items prior to the constant keywords, case insensitive.
        """
        df.loc[
            df[DESCRIPTION].str.contains("|".join(SOLVENTS), case=False),
            CATEGORY,
        ] = "solvent"

        df.loc[
            df[DESCRIPTION].str.contains("|".join(CONSUMABLES), case=False),
            CATEGORY,
        ] = "consumable"

        df.loc[
            df[DESCRIPTION].str.contains("|".join(PURIFICATION), case=False),
            CATEGORY,
        ] = "purification"

        df.loc[
            df[DESCRIPTION].str.contains("|".join(MISC), case=False),
            CATEGORY,
        ] = "miscelanous"

        df.loc[df[CATEGORY].isnull(), CATEGORY] = "other"
        return df

