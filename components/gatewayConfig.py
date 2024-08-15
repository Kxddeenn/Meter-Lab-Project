import json 

def gatewayJson(jsonFile):
    with open(jsonFile, 'r') as file:
        data = json.load(file)

    macAddress = data.get('MAC_Addresses')
    firmware = data.get('Firmware')
    modelNum = data.get('ModelNum')
    unitSerial = data.get('UnitSerial')

    modules = data.get('Modules', [])

    partModule = []
    serialModule = []

    for module in modules:
        part = module.get('Part')
        partModule.append(part)
        serial = module.get('Serial')
        serialModule.append(serial)


    return macAddress, firmware, modelNum, unitSerial, partModule, serialModule

