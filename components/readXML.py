import xml.etree.ElementTree as ET


def load6320XML(xmlFile):
    """
    Parses an XML file containing 6320 meter data and extracts MAC address, model number, and unit serial number.

    Args:
        xmlFile (str): Path to the XML file.

    Returns:
        variables: macAddress, modelNum, unitSerial
    """
    tree = ET.parse(xmlFile)
    root = tree.getroot()

    macAddress = root.find("MAC_Address").text
    modelNum = root.find("ModelNum").text
    unitSerial = root.find("UnitSerial").text

    macElem = root.find("MAC_Address")
    if macElem is not None:
        macAddress = macElem.text

    modelElem = root.find("ModelNum")
    if modelElem is not None:
        modelNum = modelElem.text

    voltageIndex = modelNum.find('V')
    if voltageIndex == -1:
        modelNum = modelNum
    
    else:
        modelNum = modelNum[:voltageIndex -4] + modelNum[voltageIndex + 1:]

    serialElem = root.find("UnitSerial")
    if serialElem is not None:
        unitSerial = serialElem.text

    # print(macAddress)
    # print(f"model number is = {modelNum}")
    # print(unitSerial)

    return macAddress, modelNum, unitSerial

def load6312XML(xmlFile):
    """
    Parses an XML file containing 6312 meter data and extracts MAC address, model number, and unit serial number.

    Args:
        xmlFile (str): Path to the XML file.

    Returns:
        variables: MAC address, model number, unit serial number, meter configuration, demand integer and pulse type.
    """
    tree = ET.parse(xmlFile)
    root = tree.getroot()


    macAddress = None
    modelNum = None
    meterConfig = None
    demandInt = None
    unitSerial = None
    pulseType = None

    macElem = root.find("MAC_Address")
    if macElem is not None:
        macAddress = macElem.text

    modelElem = root.find("ModelNum")
    if modelElem is not None:
        modelNum = modelElem.text

    voltageIndex = modelNum.find('V')
    if voltageIndex == -1:
        modelNum = modelNum
    
    else:
        modelNum = modelNum[:voltageIndex -4] + modelNum[voltageIndex + 1:]

    meterElem = root.find("MeterConfig")
    if meterElem is not None:
        meterConfig = meterElem.text

    demandElem = root.find("Demand_Interval")
    if demandElem is not None:
        demandInt = demandElem.text

    serialElem = root.find("UnitSerial")
    if serialElem is not None:
        unitSerial = serialElem.text

    pulseElem = root.find("Pulse_Type")
    if pulseElem is not None:
        pulseType = pulseElem.text

    return macAddress, modelNum, unitSerial, meterConfig, demandInt, pulseType


def load6303XML(xmlFile):
    """
    Parses an XML file containing 6303 meter data and extracts MAC address, model number, and unit serial number.

    Args:
        xmlFile (str): Path to the XML file.

    Returns:
        tuple: Tuple containing MAC address, model number , unit serial number, meter configuration, demand integer and pulse type.
    """
    tree = ET.parse(xmlFile)
    root = tree.getroot()

    macAddress = None
    modelNum = None
    meterConfig = None
    demandInt = None
    unitSerial = None
    pulseType = None

    macAddress = root.find("MAC_Address").text
    modelNum = root.find("ModelNum").text
    meterConfig = root.find("MeterConfig").text
    demandInt = root.find("Demand_Interval").text
    unitSerial = root.find("UnitSerial").text
    pulseType = root.find("Pulse_Type").text

    macElem = root.find("MAC_Address")
    if macElem is not None:
        macAddress = macElem.text

    modelElem = root.find("ModelNum")
    if modelElem is not None:
        modelNum = modelElem.text

    voltageIndex = modelNum.find('V')
    if voltageIndex == -1:
        modelNum = modelNum
    
    else:
        modelNum = modelNum[:voltageIndex -4] + modelNum[voltageIndex + 1:]

    meterElem = root.find("MeterConfig")
    if meterElem is not None:
        meterConfig = meterElem.text

    demandElem = root.find("Demand_Interval")
    if demandElem is not None:
        demandInt = demandElem.text

    serialElem = root.find("UnitSerial")
    if serialElem is not None:
        unitSerial = serialElem.text

    pulseElem = root.find("Pulse_Type")
    if pulseElem is not None:
        pulseType = pulseElem.text

    
    return macAddress, modelNum, unitSerial, meterConfig, demandInt, pulseType


def findProduct(xmlFile, product):
    """
    Parses an XML file to determine the specific product type based on the 'ModelNum' element.

    Args:
        xmlFile (str): Path to the XML file.
        product (str): Product identifier ('6312', '6320', '6303').

    Returns:
        str: Product type ('6312-2P', '6312-1P', '6312-3P', '6320-2P', '6320-1P', '6320-3P', '6303-2P', '6303-1P', '6303-3P').
             Returns None if the product type cannot be determined.
    """
    tree = ET.parse(xmlFile)
    root = tree.getroot()

    modelNum = root.find("ModelNum").text

    if product == "6312":
        if "2P" in modelNum:
            return "6312-2P"

        elif "1P" in modelNum:
            return "6312-1P"

        elif "3P" in modelNum:
            return "6312-3P"
        
    elif product == "6320":
        if "2P" in modelNum:
            return "6320-2P"

        elif "1P" in modelNum:
            return "6320-1P"

        elif "3P" in modelNum:
            return "6320-3P"
    
    elif product == "6303":
        if "2P" in modelNum:
            return "6303-2P"

        elif "1P" in modelNum:
            return "6303-1P"

        elif "3P" in modelNum:
            return "6303-3P"        

        
