# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import random
import xlwt

######################################        the path for storing the dataset       ##################################
mDir = 'D:\\yzwang\\PycharmProjects\\GitHubSH\\'


########################################################################################################################
################################################# data Generator #######################################################
########################################################################################################################
def dataGeneration(AQC, Block, Container, DIS_AQC_Block, SEED, MAXNUM, ManOrRand, BlockNumMan):

    Loc_O = (0, 0)  # Position coordinate origin (x, y)=(0, 0)

    #***********************  Coordinates of the Quay Crane(QC) and the Blocks     ************************************#
    AQCSet = {}     # the set of QCs, i.e., AQCSet = {1:(0,100),...}
    BlockSet = {}   # the set of blocks, i,e, BlockSet = {1:(140, 33),...}
    for i in range(1, AQC['NUM']+1):
        AQCSet[i] = (Loc_O[0], AQC['DIS'] * i)
    for i in range(1, Block['NUM'] + 1):
        BlockSet[i] = (DIS_AQC_Block, (Block['CON'] * Container['Wdh'] + Block['DIS']) * int(i))
    # print(AQCSet)
    # print(BlockSet)


    #******************  Specify the number of tasks per box area - given manually or randomly generated *************#
    random.seed(SEED)       # a random seed
    JobNumBlockID = {}      # the number of tasks in each Block, i,e, {Block: number of tasks}
    if ManOrRand == 0:      # given manually
        JobNumBlockID = BlockNumMan
    elif ManOrRand == 1:    # randomly generated
        for k in BlockSet.keys():
            JobNumBlockID[k] = random.randint(0, MAXNUM)
    # print(JobNumBlockID)


    #*********************************** Generate the tasks in Blocks ***********************************************#
    JOBBlock = {}
    for k in JobNumBlockID.keys():
        temp_list = []
        for i in range(JobNumBlockID[k]):
            temp_gendata = [random.randint(1, Block['Len']), random.randint(1, Block['CON'])]
            if not temp_gendata in temp_list:       # Remove the duplicate data
                temp_list.append(temp_gendata)
        JOBBlock[k] = temp_list
    # There is no duplicate data of tasks in the block
    for k in JOBBlock.keys():
        JobNumBlockID[k] = len(JOBBlock[k])
    # print(JobNumBlockID)


    #******************************** Calculate the coordinates corresponding to the task    **************************#
    JOBcoordinate = {}      # JOBcoordinate = {1: [[440, 54], [176, 48]], 2: [[512, 84]],...}
    for k in JOBBlock.keys():
        temp_coord = []
        for v in JOBBlock[k]:
            temp_coord.append([BlockSet[k][0] + v[0] * Container['Len'],
                               BlockSet[k][1] + v[1] * Container['Wdh']])
        JOBcoordinate[k] = temp_coord
    # print(JOBcoordinate)

    # print('Coordinates of QCs:', AQCSet)
    # print('Coordinates of Blocks:', BlockSet)
    # print('the number of tasks:', JobNumBlockID)
    # print('the tasks in the block: ', JOBBlock)
    # print('Coordinates of task in the block: ', JOBcoordinate)

    return AQCSet, JobNumBlockID, JOBBlock, JOBcoordinate, BlockSet


#####################################  data export to EXCEL

########################################################################################################################
################################################# data export to EXCEL #################################################
########################################################################################################################
def data2Excel(AQC_coord, Job_N, BlockBay, Block_coord, ManOrAuto, CombMan):

    #**************************** QC matches Blocks: Let the Blocks select AQC    ************************************#
    BN = list(BlockBay.keys())
    AN = list(AQC_coord.keys())
    Comb = {}                # Comb = {Block: AQC}
    if ManOrAuto == 0:       # Manual Matching (given)
        Comb = CombMan
    elif ManOrRand == 1:     # Random Matching (Not given)
        for i in BN:
            Comb[i] = random.randint(1, len(AN))

    N = np.sum(list(Job_N.values()))  # the number of tasks

    #**********************************************    Declare the Excel header    ************************************#
    workbook = xlwt.Workbook()
    worksheet1 = workbook.add_sheet('task')
    worksheet1.write(0, 0, 'No.')
    worksheet1.write(0, 1, 'Type')
    worksheet1.write(0, 2, 'AQC_ID')
    worksheet1.write(0, 3, 'AQC_X')
    worksheet1.write(0, 4, 'AQC_Y')
    worksheet1.write(0, 5, 'Block_ID')
    worksheet1.write(0, 6, 'Block_Bay')
    worksheet1.write(0, 7, 'Block_Row')
    worksheet1.write(0, 8, 'Yard_X')
    worksheet1.write(0, 9, 'Yard_Y')

    #**********************************  Write AQC coordinates, number of tasks, etc.  ********************************#
    idx = 1
    for i in Comb.keys():
        inx = 0
        for v in Block_coord[i]:
            worksheet1.write(idx, 0, idx)                     # N
            worksheet1.write(idx, 1, random.randint(0, 1))    # the Type of tasks: 1-import container; 0-export container
            worksheet1.write(idx, 2, Comb[i])                 # AQC_ID
            worksheet1.write(idx, 3, AQC_coord[Comb[i]][0])   # AQC_X
            worksheet1.write(idx, 4, AQC_coord[Comb[i]][1])   # AQC_Y
            worksheet1.write(idx, 5, int(i))                  # Block_ID
            worksheet1.write(idx, 6, BlockBay[i][inx][0])     # Block_Bay
            worksheet1.write(idx, 7, BlockBay[i][inx][0])     # Block_Row
            worksheet1.write(idx, 8, v[0])                    # Yard_X
            worksheet1.write(idx, 9, v[1])                    # Yard_Y
            inx += 1
            idx += 1

    fnSolution = mDir
    fnName = 'N' + str(N) + '_AQC' + str(len(AQC_coord)) + '_Block' + str(len(BlockBay)) + '.xls'
    workbook.save(fnSolution + fnName)

    return fnName

def readExcel(fnName, AGV, ASC, BlockSet, N, NK, NA, NB):
    dataset = pd.read_excel(mDir + fnName, sheet_name='task', skiprows=0)
    df = pd.DataFrame(dataset)
    dataExp = df.loc[0: 5, ['No.', 'Type', 'AQC_ID', 'AQC_X', 'AQC_Y', 'Block_ID', 'Block_Bay', 'Block_Row', 'Yard_X', 'Yard_Y']]
    # print(dataExp)

    #*************************************************   various sets  ************************************************#
    J = list(df['No.'].values)[0: N]    # task: 1,2,...
    TP = list(df['Type'].values)[0: N]
    AQC_ID = [0] + list(df['AQC_ID'].values)[0: N]
    X1 = [0] + list(df['AQC_X'].values)[0: N]
    Y1 = [0] + list(df['AQC_Y'].values)[0: N]
    B_ID = [0] + list(df['Block_ID'].values)[0: N]
    B_Bay = [0] + list(df['Block_Bay'].values)[0: N]
    B_Row = [0] + list(df['Block_Row'].values)[0: N]
    X2 = [0] + list(df['Yard_X'].values)[0: N]
    Y2 = [0] + list(df['Yard_Y'].values)[0: N]

    JS = list(range(0, N + 1))  # the set of tasks with dummy starting task
    JE = list(range(1, N + 2))  # the set of tasks with dummy finishing task
    JD = list(range(0, N + 2))  # the set of tasks with dummy starting and finishing task
    JI = [i + 1 for i, x in enumerate(TP) if x == 1]  # set of import container tasks
    JO = [j + 1 for j, x in enumerate(TP) if x == 0]  # set of export container tasks
    print(JI, JO)

    K = set(list(range(1, NK + 1)))     # the set of AGVs
    A = set(list(range(1, NA + 1)))     # set of direct transferring site, i.e., pads
    B = set(list(range(1, NB + 1)))     # set of buffer transferring site, i.e., AGV partners

    # the coordinates of BlockSet: {1: (140, 33), 2: (140, 66), 3: (140, 99)}
    #*********************************     the loading duration of AGV or ASC   ***************************************#
    TTK, TTY= np.zeros([N + 1, 1]), np.zeros([N + 1, 1])
    for i in J:
        TTK[i][0] = (1 / AGV['SPD']) * \
                    (abs(X1[i] - BlockSet[B_ID[i]][0]) + abs(Y1[i] - BlockSet[B_ID[i]][1]))
        TTY[i][0] = (1 / ASC['SPD']) * \
                    (abs(X2[i] - BlockSet[B_ID[i]][0]) + abs(Y2[i] - BlockSet[B_ID[i]][1]))

    #*********************************     the travel time of AGV or ASC   ********************************************#
    TK, TY = np.zeros([N + 2, N + 2]), np.zeros([N + 2, N + 2])
    for a in range(N + 2):
        for b in range(N + 2):
            if a in JI:
                if b in JI:
                    TK[a][b] = (1 / AGV['SPD']) * (abs(BlockSet[B_ID[a]][0] - X1[b]) + abs(BlockSet[B_ID[a]][1] - Y1[b]))
                    TY[a][b] = (1 / ASC['SPD']) * (abs(BlockSet[B_ID[a]][0] - X2[b]) + abs(BlockSet[B_ID[a]][1] - Y2[b]))
                elif b in JO:
                    TK[a][b] = (1 / AGV['SPD']) * (abs(BlockSet[B_ID[a]][0] - BlockSet[B_ID[b]][0]) +
                                                   abs(BlockSet[B_ID[a]][1] - BlockSet[B_ID[b]][1]))
                    TY[a][b] = (1 / ASC['SPD']) * (abs(X2[a] - X2[b]) + abs(Y2[a] - Y2[b]))
            elif a in JO:
                if b in JI:
                    TK[a][b] = (1 / AGV['SPD']) * (abs(X1[a] - X1[b]) + abs(Y1[a] - Y1[b]))
                    TY[a][b] = (1 / ASC['SPD']) * (abs(BlockSet[B_ID[a]][0] - BlockSet[B_ID[b]][0]) +
                                                   abs(BlockSet[B_ID[a]][1] - BlockSet[B_ID[b]][1]))
                elif b in JO:
                    TK[a][b] = (1 / AGV['SPD']) * (abs(BlockSet[B_ID[a]][0] - X1[b]) + abs(BlockSet[B_ID[a]][1] - Y1[b]))
                    TY[a][b] = (1 / ASC['SPD']) * (abs(BlockSet[B_ID[a]][0] - X2[b]) + abs(BlockSet[B_ID[a]][0] - Y2[b]))

    for a in range(N + 2):
        for b in range(N + 2):
            if a == 0:
                TK[a][b] = 0
                TY[a][b] = 0
            if b == 0:
                TK[a][b] = 9999
                TY[a][b] = 9999
            if b == N + 1:
                TK[a][b] = 0
                TY[a][b] = 0
            if a == N + 1:
                TK[a][b] = 9999
                TY[a][b] = 9999
    TK[0][N + 1] = 9999
    TY[0][N + 1] = 9999

    return J,  K, A, B, TTK, TTY, TK, TY

########################################################################################################################
#******************************************* Declare variables and assign values   ************************************#
########################################################################################################################
AQC = {'NUM': 3, 'DIS': 100}            # QC parameters: number, the distance between two QCs(m)
AGV = {'SPD': 6, 'BK': 10, 'TQ': 15}    # AGV parameters：operational speed(m/s), lift/drop time for a container(s), handling time under QCs(s)
ASC = {'SPD': 4, 'BY': 5, 'TY': 10}     # ASC parameters：operational speed(m/s), the time of capturing the releasing a container(s),lift/drop time for a container(s)


#*******************************************  parameters of Blocks     ************************************************#
#  the number of containers, the interval between two adjacent containers (m), the maximum number of bays,
#  the number of containers with the same bays, the number of pads, the number of AGV partners
Block = {'NUM': 3, 'DIS': 3, 'Len': 40, 'CON': 10, 'DH': 3, 'BH': 3}
Container = {'Len': 12, 'Wdh': 3}       # the length of a container: len(m)、wdh(m)
DIS_AQC_Block = 140                     # the vertical distance from quayside bridge to yard (m)

SEED = 1                                # the random seed
MAXNUM = 50                             # the maximum number of tasks in the Bolck
ManOrRand = 0                           # 0: Manual; 1: random
BlockNumMan = {1: 5, 2: 10, 3: 10}      # manual seeting. {Block: number of tasks}
#################################################################################################
AQCSet, JobNumBlockID, JOBBlock, JOBcoordinate, BlockSet = dataGeneration(AQC, Block, Container,
                                                                          DIS_AQC_Block, SEED,
                                                                          MAXNUM, ManOrRand, BlockNumMan)

ManOrAuto = 0                   # mathcing the Block_ID and AQC_ID. the pattern: 0-Manual matching, 1-Random matching.
CombMan = {1: 3, 2: 2, 3: 1}    # {Bolck_ID: AQC_ID}
fnName = data2Excel(AQCSet, JobNumBlockID, JOBBlock, JOBcoordinate, ManOrAuto, CombMan)

N = 10  # the number of tasks
NK = 3  # the number of AGV
NA = 3  # the number of pads
NB = 3  # the number of AGV partner
J,  K, A, B, TTK, TTY, TK, TY = readExcel(fnName, AGV, ASC, BlockSet, N, NK, NA, NB)
