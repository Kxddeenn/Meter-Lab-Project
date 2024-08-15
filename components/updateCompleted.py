from openpyxl import load_workbook
import matplotlib.pyplot as plt
import numpy as np



def refreshCompleted():

    sealingLog = load_workbook(r'excelFiles\SealingLog.xlsx', data_only=True)
    graphSheet = sealingLog['2024Data']

    columns = ['B', 'C']

    data = []
    for col in columns:
        columnData = [graphSheet[f'{col}{row}'].value for row in range(2, 15)]
        data.append(columnData)

    dataArray = np.array(data)
    dataArray = dataArray[:, ::-1]
   
    headers = dataArray[0]
    values = dataArray[1]

    dataDict = [{'Month': headers[i], 'Value': int(values[i])} for i in range(len(headers))]

    fig, ax = plt.subplots(figsize=(8, 6)) 

    cols = 2
    rows = 14
    ax.set_ylim(0, rows - 0.5)
    ax.set_xlim(-0.5, cols - 0.5)
    ax.axis('off')


    for idx, d in enumerate(dataDict):
        fontweight = 'bold' if idx == len(dataArray) - 2 else 'normal'
        fontsize = 12 if idx == len(dataArray) - 2 else 10
        ax.text(x=0, y=idx, s=d['Month'], va='center', ha='left', fontsize=fontsize, fontweight=fontweight)
        ax.text(x=1, y=idx, s=d['Value'], va='center', ha='right', fontsize=fontsize, fontweight=fontweight)
    
    ax.text(0,13, "Month", weight="bold", fontsize=12, ha='left')
    ax.text(0.85,13, "Number", weight="bold", fontsize=12 , ha='left')

    for row in range(rows):
        ax.plot(
            [-0.5, cols - 0.5],
            [row -.5, row - .5],
            ls=':',
            lw='.5',
            c='grey'
        )

    ax.set_title(
	'Meter Seals Completed',
	loc='center',
	fontsize=18,
	weight='bold',
    pad=20
)
    plt.savefig('images\meterseals.png', bbox_inches='tight', pad_inches=0.1)
    plt.close()

    sealingLog.close()
