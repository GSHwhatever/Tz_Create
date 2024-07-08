from faker import Faker

phone_lis = []
fake = Faker('zh_CN')
for i in range(100):
    phone_lis.append(fake.phone_number())

phone_set = set(phone_lis)
print(phone_set)
print(len(phone_set))
"""
{'15335806901', '15807712890', '13365691546', '14737213533', '14744169494', '13872174566', '13501487624', '18880447363', '13197257441', '13682647331', '14587123701', '13518012016', '15931436317', '13839856133', '13245634025', '13805051169', '14722508132', '14572807108', '18159779880', '13421288686', '18721210550', '13419656316', '13303384200', '13421988352', '14528609545', '18047764017', '14545502757', '15630795670', '14515789045', '13689506393', '18980444868', '15616449340', '13699057347', '18756894866', '13771606922', '13869888199', '13861104222', '15299520712', '18224105233', '13217641812', '18751585980', '13273234363', '14572130934', '13195577767', '13902458873', '15146460808', '13904630892', '18916035208', '15716603848', '13321100939', '18104877540', '13527655258', '15901832645', '15813169140', '13759478609', '13652189664', '15965634560', '13162080202', '18840890401', '18657577170', '14588787828', '18514131676', '18707109197', '18665368641', '15104865381', '18046860759', '15231298712', '18711136211', '14568929768', '18129979157', '18121679874', '13613712569', '15357435439', '13411476045', '15556191668', '18576797768', '18516991301', '18687345412', '13681375934', '13500274379', '18274880183', '13208213176', '13772604933', '18178086615', '15580620695', '15792894378', '13853777335', '15187856821', '15577911106', '15383204459', '13495797296', '18071821086', '13990539789', '18648677676', '14766647549', '18090528239', '14592734318', '18090716740', '15765882933', '15907753383'}
"""