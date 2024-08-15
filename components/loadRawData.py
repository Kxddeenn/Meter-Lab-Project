import pandas


def raw6312(csvFile):
    """
    Reads specific columns from a CSV file into lists only for 6312 meters.

    Args:
        csvFile (str): Path to the CSV file.

    Returns:
        tuple: Tuple containing lists for each column.
    """
    dataFrame = pandas.read_csv(csvFile)

    sequenceValue = []
    meterID = []
    serialNum = []
    modelType = []
    leftCoil = []
    rightCoil = []
    PFValue = []
    FLValue = []
    LLValue = []

    sequenceValue = dataFrame.iloc[0:13, 0].tolist()
    meterID = dataFrame.iloc[0:13, 1].tolist()
    serialNum = dataFrame.iloc[0:13, 2].tolist()
    modelType = dataFrame.iloc[0:13, 3].tolist()
    leftCoil = dataFrame.iloc[0:13, 4].tolist()
    rightCoil = dataFrame.iloc[0:13, 5].tolist()
    PFValue = dataFrame.iloc[0:13, 6].tolist()
    FLValue = dataFrame.iloc[0:13, 7].tolist()
    LLValue = dataFrame.iloc[0:13, 8].tolist()

    return (
        sequenceValue,
        meterID,
        serialNum,
        modelType,
        leftCoil,
        rightCoil,
        PFValue,
        FLValue,
        LLValue,
    )

def raw63123(csvFile):
    """
    Reads specific columns from a CSV file into lists only for 6312-3P meters.

    Args:
        csvFile (str): Path to the CSV file.

    Returns:
        tuple: Tuple containing lists for each column.
    """
    dataFrame = pandas.read_csv(csvFile)

    sequenceValue = []
    meterID = []
    serialNum = []
    modelType = []
    leftCoil = []
    rightCoil = []
    middleCoil = []
    PFValue = []
    FLValue = []
    LLValue = []

    sequenceValue = dataFrame.iloc[0:8, 0].tolist()
    meterID = dataFrame.iloc[0:8, 1].tolist()
    serialNum = dataFrame.iloc[0:8, 2].tolist()
    modelType = dataFrame.iloc[0:8, 3].tolist()
    leftCoil = dataFrame.iloc[0:8, 4].tolist()
    middleCoil = dataFrame.iloc[0:8, 5].tolist()
    rightCoil = dataFrame.iloc[0:8, 6].tolist()
    PFValue = dataFrame.iloc[0:8, 7].tolist()
    FLValue = dataFrame.iloc[0:8, 8].tolist()
    LLValue = dataFrame.iloc[0:8, 9].tolist()

    return (
        sequenceValue,
        meterID,
        serialNum,
        modelType,
        leftCoil,
        rightCoil,
        PFValue,
        FLValue,
        LLValue,
        middleCoil
    )

def raw6303(csvFile):
    """
    Reads specific columns from a CSV file into lists only for 6303 meters.

    Args:
        csvFile (str): Path to the CSV file.

    Returns:
        tuple: Tuple containing lists for each column.
    """
    dataFrame = pandas.read_csv(csvFile)

    sequenceValue = []
    meterID = []
    serialNum = []
    modelType = []
    leftCoil = []
    rightCoil = []
    PFValue = []
    FLValue = []
    LLValue = []

    sequenceValue = dataFrame.iloc[0:4, 0].tolist()
    meterID = dataFrame.iloc[0:4, 1].tolist()
    serialNum = dataFrame.iloc[0:4, 2].tolist()
    modelType = dataFrame.iloc[0:4, 3].tolist()
    leftCoil = dataFrame.iloc[0:4, 4].tolist()
    rightCoil = dataFrame.iloc[0:4, 5].tolist()
    PFValue = dataFrame.iloc[0:4, 6].tolist()
    FLValue = dataFrame.iloc[0:4, 7].tolist()
    LLValue = dataFrame.iloc[0:4, 8].tolist()

    return (
        sequenceValue,
        meterID,
        serialNum,
        modelType,
        leftCoil,
        rightCoil,
        PFValue,
        FLValue,
        LLValue,
    )

def raw6320(csvFile):
    """
    Reads specific columns from a CSV file into lists only for 6320 meters.

    Args:
        csvFile (str): Path to the CSV file.

    Returns:
        tuple: Tuple containing lists for each column.
    """
    dataFrame = pandas.read_csv(csvFile)

    sequenceValue = []
    meterID = []
    serialNum = []
    modelType = []
    leftCoil = []
    rightCoil = []
    PFValue = []
    FLValue = []
    LLValue = []
    SeriesValue = []

    sequenceValue = dataFrame.iloc[0:20, 0].tolist()
    meterID = dataFrame.iloc[0:20, 1].tolist()
    serialNum = dataFrame.iloc[0:20, 2].tolist()
    modelType = dataFrame.iloc[0:20, 3].tolist()
    leftCoil = dataFrame.iloc[0:20, 4].tolist()
    rightCoil = dataFrame.iloc[0:20, 5].tolist()
    LLValue = dataFrame.iloc[0:20, 9].tolist()
    PFValue = dataFrame.iloc[0:20, 6].tolist()
    FLValue = dataFrame.iloc[0:20, 8].tolist()
    SeriesValue = dataFrame.iloc[0:20, 7].tolist()

    return (
        sequenceValue,
        meterID,
        serialNum,
        modelType,
        leftCoil,
        rightCoil,
        PFValue,
        FLValue,
        LLValue,
        SeriesValue,
    )

