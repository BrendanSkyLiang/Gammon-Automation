from plxscripting.easy import *
import numpy

def listToString(s): 
    
    # initialize an empty string
    str1 = "" 
    
    # traverse in the string  
    for ele in s: 
        str1 += ele  
    
    # return string  
    return str1 
    
    
#input parameters
localport = input('input local host port: ')
#localport = 10000
serverPassword = input('input local host pasword: ')
#serverPassword = 'QV^g<Nr9vCcn%r44'
filePathLocation = input('enter file save path: ')

#Data Pulling from PLAXIS

s_o, g_o = new_server('localhost', localport, password = serverPassword)

stage_number = 0
AR = [('Stage Number', 'Stage Name', 'xcood','ycoord','force')]

for phase in g_o.Phases[1:]:
    stage_number = stage_number + 1
    try:
        anchorF = g_o.getresults(phase, g_o.ResultTypes.FixedEndAnchor.AnchorForce2D, 'node') 
        anchorY = g_o.getresults(phase, g_o.ResultTypes.FixedEndAnchor.Y, 'node') 
        anchorX = g_o.getresults(phase, g_o.ResultTypes.FixedEndAnchor.X,'node')
        anchorPhase = listToString(g_o.echo(phase.Identification))
        N = [anchorPhase]*len(anchorX)
        P = [stage_number]*len(anchorX)
        x = zip(P, N, anchorX,anchorY,anchorF)
        AR.append(tuple(x))
        print("Completed {}".format(phase.Name))
    except:
        print("No Active Anchor in {}".format(phase.Name))
        
#Data Organisation and Export to CSV

#dataArray = []
#for a in range(1,len(AR)):
    #for b in range(0,len(AR[a])):
        #dataArray.append(AR[a][b])
        
#export = numpy.array(dataArray)


#numpy.savetxt("C:/Users/brendanlia/Desktop/new2.csv", export, delimiter = ',', fmt=['%s', '%s', '%s', '%s', '%s'], header = 'Stage Number, Phase Name, xcoord, ycoord, force (kN)')


AR1 = AR[1:]

compiledList = []

passThroughList = []

for c in range(0,len(AR1)):
    if len(AR1[c]) >= 1:
        passThroughList = [AR1[c][0][0], AR1[c][0][1]]
        for d in range(0,len(AR1[c])):
            passThroughList.append(AR1[c][d][2])
            passThroughList.append(AR1[c][d][3])
            passThroughList.append(AR1[c][d][4])
        compiledList.append(passThroughList)
    else:
        print('error')
        


alpha = []
for x in range(0,len(compiledList)):
    alpha.append(compiledList[x])
    
final = numpy.array(alpha,dtype = object)

numpy.savetxt(filePathLocation, final, delimiter=',', fmt = '%s',header='Stage Number, Phase Name, xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN), xcoord, ycoord, force (kN)')

input('Press ENTER to exit')



