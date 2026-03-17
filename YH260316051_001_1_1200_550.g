N01  P3000 = 1200[cutsizeX+30]
N02  P3001 = 550[cutsizeY+30]
N03  P3002 = 0
N04  P3003 = 0
N05  P3004 = 0
N06  P3005 = 0
N07  P3006 = 1[默认]
N08  P3007 = 5mm超白[material_code]
N09  P3008 = 1[默认]
N10  P3009 = 1[默认]
N11  P3010 = 
N12  P3011 = 5[thickness]
N13  P4001= 1145_499_0_0_1143_497_xyg-F_1_中二补_D260311131_5112405911____5112405911__497x1143/A_497_中二补_ _3 1_________

N13  P4001= {cutsizeX}_{cutsizeY}_0_0_{displayX}_{displayY}_{customer_name}_1_{group_number}_{order_number}_{dm_code}____{dm_code}__{order_size}_{reference_edge}_{group_number}_{code_3c_position}_{dm_code_position}_________


G17
G92 X0 Y0
G90
G00 X3 Y{cutsizeY}
M03
M09
G01 X{cutsizeX} Y{cutsizeY}
M10
G00 X{cutsizeX} Y3
M09
G01 X{cutsizeX} Y[{cutsizeY}+30-3]
M10
M04
G90G00X0Y0Z0
M23
M24
M30
