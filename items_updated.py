import pandas as pd
import sys

# Create constants to work with the dataframe consistently
CODE = "code"
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

MISC = ["SABLE", "SILICE", "GRANU", "SODIUM", "JAVEL", "ACI"]


class Items:
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

        # create a variable giving all categories for the stock
        self.categories = [i for i in self.df["category"].unique()]

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
                    df.columns[2]: DESCRIPTION,
                },
                inplace=True,
            )
        except IndexError:
            print("Wrong dataframe format")
            sys.exit(1)

        # Homogenize data strings in description column
        df.loc[:, DESCRIPTION] = df[DESCRIPTION].str.replace("\n", " ")
        df.loc[:, DESCRIPTION] = df[DESCRIPTION].str.capitalize()
        df.loc[:, DESCRIPTION] = df[DESCRIPTION].str.strip()
        
        # sort data by description
        df.sort_values(by=DESCRIPTION, inplace=True)
        
        return df

    @classmethod
    def __categorize_items(cls, df: pd.DataFrame) -> pd.DataFrame:
        """
        Input a dataframe containing all items and create a columns "category" to categorize all items prior to the constant keywords, case insensitive.
        Update the self.categories variable
        """
        df.loc[
            df[DESCRIPTION].str.contains("|".join(SOLVENTS), case=False),
            CATEGORY,
        ] = "solvents"

        df.loc[
            df[DESCRIPTION].str.contains("|".join(CONSUMABLES), case=False),
            CATEGORY,
        ] = "consumables"

        df.loc[
            df[DESCRIPTION].str.contains("|".join(PURIFICATION), case=False),
            CATEGORY,
        ] = "purification"

        df.loc[
            df[DESCRIPTION].str.contains("|".join(MISC), case=False),
            CATEGORY,
        ] = "miscelanous"

        df.loc[df[CATEGORY].isnull(), CATEGORY] = "others"
        return df

    def select_category(self, category: str) -> pd.DataFrame:
        """
        From a category, select the corresponding items from the dataframe with their codes
        """
        return self.df.loc[self.df[CATEGORY] == category, [DESCRIPTION, CODE]]

    def item_from_code(self, code: int) -> str:
        """
        return the item's name from his code
        """
        return self.df.loc[self.df[CODE] == code, DESCRIPTION].iloc[0]

    def __str__(self):
        return str(self.df)
