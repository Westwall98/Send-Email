import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

# body = BeautifulSoup(body,"lxml").select("strong")[1].text.replace('%%FirstName%%','Ethan')
# print(body)

userlist = [{'User': 'Aaron Ren (ext.)', 'User First Name': 'Aaron', 'Email': 'Aaron.Ren@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ada Liu (ext.)', 'User First Name': 'Ada', 'Email': 'Ada.LIU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Adrian Pan (ext.)', 'User First Name': 'Adrian', 'Email': 'Adrian.Pan-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ajuan MENG (ext.)', 'User First Name': 'Ajuan', 'Email': 'Ajuan.MENG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Alan LIU (dummy)', 'User First Name': 'Alan', 'Email': 'Alan.LIUS@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Albert Chen', 'User First Name': 'Albert', 'Email': 'Albert.CHEN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Alex Liang (ext.)', 'User First Name': 'Alex', 'Email': 'alex.liang-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Alexis YU', 'User First Name': 'Alexis', 'Email': 'Alexis.YU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Alice Zhang (FN)', 'User First Name': 'Alice', 'Email': 'Alice.ZHANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Allen Hong', 'User First Name': 'Allen', 'Email': 'Allen.HONG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Alvin Zhang', 'User First Name': 'Alvin', 'Email': 'alvin.zhang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Amanda Song', 'User First Name': 'Amanda', 'Email': 'amanda.song@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Amei Li', 'User First Name': 'Amei', 'Email': 'Amei.Li@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Amy WANG (ext.)', 'User First Name': 'Amy', 'Email': 'amy.wang-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Andrew TANG', 'User First Name': 'Andrew', 'Email': 'andrew.tang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Angela Chen', 'User First Name': 'Angela', 'Email': 'angela.chen@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Angela XUS (Dummy)', 'User First Name': 'Angela', 'Email': 'Angela.XUS@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Annie Gong', 'User First Name': 'Annie', 'Email': 'Annie.GONG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Anthony Chen', 'User First Name': 'Anthony', 'Email': 'anthony.chen@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'April Wu', 'User First Name': 'April', 'Email': 'april.wu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Asa HE (ext.)', 'User First Name': 'Asa', 'Email': 'asa.he-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Balton Luo', 'User First Name': 'Balton', 'Email': 'Balton.LUO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Bella Ding', 'User First Name': 'Bella', 'Email': 'bella.ding1@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ben Wang', 'User First Name': 'Ben', 'Email': 'ben.wang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Benic Zheng', 'User First Name': 'Benic', 'Email': 'Benic.ZHENG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Bin DU (ext.)', 'User First Name': 'Bin', 'Email': 'Bin.DU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Bin ZHOU', 'User First Name': 'Bin', 'Email': 'Bin.ZHOU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Bo Yang', 'User First Name': 'Bo', 'Email': 'Bo.YANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Bonnie Xu', 'User First Name': 'Bonnie', 'Email': 'bonnie.xu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Brian Cao', 'User First Name': 'Brian', 'Email': 'brian.cao@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Bruce Liang', 'User First Name': 'Bruce', 'Email': 'bruce.liang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Byron Zhu', 'User First Name': 'Byron', 'Email': 'Byron.ZHU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Cara Ding', 'User First Name': 'Cara', 'Email': 'cara.ding@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Carrie Cao', 'User First Name': 'Carrie', 'Email': 'carrie.cao@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Cassie Li', 'User First Name': 'Cassie', 'Email': 'Yiwei.LI@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Cathy Ma', 'User First Name': 'Cathy', 'Email': 'Cathy.MA@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Chamber Li (ext.)', 'User First Name': 'Chamber', 'Email': 'Chamber.LI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Chaoqun Wu (ext.)', 'User First Name': 'Chaoqun', 'Email': 'chaoqun.wu-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Charly Mao', 'User First Name': 'Charly', 'Email': 'Charly.MAO@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Chengyue Zhang (ext.)', 'User First Name': 'Chengyue', 'Email': 'Chengyue.ZHANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'China Compliance Officer', 'User First Name': 'China', 'Email': 'China-compliance-officer@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Chloe Xu', 'User First Name': 'Chloe', 'Email': 'Chloe.XU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Chris YANG', 'User First Name': 'Chris', 'Email': 'Chris.Yang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Chunhua LIU (ext.)', 'User First Name': 'Chunhua', 'Email': 'Chunhua.Liu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Cici Xi', 'User First Name': 'Cici', 'Email': 'cici.xi@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Civet Hu', 'User First Name': 'Civet', 'Email': 'civet.hu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Claude Liu (ext.)', 'User First Name': 'Claude', 'Email': 'claude.liu-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Colin Yang (ext.)', 'User First Name': 'Colin', 'Email': 'colin.yang-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Connie Hu', 'User First Name': 'Connie', 'Email': 'connie.hu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Cudi Chen', 'User First Name': 'Cudi', 'Email': 'cudi.chen@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Daiqing Chao', 'User First Name': 'Daiqing', 'Email': 'Daiqing.CHAO@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Daley Du (ext.)', 'User First Name': 'Daley', 'Email': 'daley.du-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Daniel YUAN', 'User First Name': 'Daniel', 'Email': 'Daniel.YUAN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'David Chen', 'User First Name': 'David', 'Email': 'David.Chen@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'David Sun (ext.)', 'User First Name': 'David', 'Email': 'David.Sun-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'David Zhong', 'User First Name': 'David', 'Email': 'David.ZHONG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Demi Liu', 'User First Name': 'Demi', 'Email': 'demi.liu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Dev01 PRC-ext', 'User First Name': 'Dev01', 'Email': 'Dev01.PRC-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Dirk Xu', 'User First Name': 'Dirk', 'Email': 'dirk.xu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Dong YAO (ext.)', 'User First Name': 'Dong', 'Email': 'Dong.YAO@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Doris Deng', 'User First Name': 'Doris', 'Email': 'Doris.DENG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Dylan Liu', 'User First Name': 'Dylan', 'Email': 'dylan.liu@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Echo Xie', 'User First Name': 'Echo', 'Email': 'Echo.XIE@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Elaine Zhou', 'User First Name': 'Elaine', 'Email': 'elaine.zhou@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ella ZHOU', 'User First Name': 'Ella', 'Email': 'ella.zhou@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Elmo Zeng (ext.)', 'User First Name': 'Elmo', 'Email': 'elmo.zeng-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Emily Chen (ext.)', 'User First Name': 'Emily', 'Email': 'emily.chen-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Emol Fan', 'User First Name': 'Emol', 'Email': 'Emol.FAN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Eric Wang', 'User First Name': 'Eric', 'Email': 'eric.wang1@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Ethan Cui (ext.)', 'User First Name': 'Ethan', 'Email': 'ethan.cui-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Even SUN (ext.)', 'User First Name': 'Even', 'Email': 'Even.SUN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Fanny CHEN (ext.)', 'User First Name': 'Fanny', 'Email': 'Fanny.CHEN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Fei Guo', 'User First Name': 'Fei', 'Email': 'Fei.GUO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Fengning Jiang', 'User First Name': 'Fengning', 'Email': 'Fengning.JIANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Fiona WANG (ext.)', 'User First Name': 'Fiona', 'Email': 'fiona.wang-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Flora LI', 'User First Name': 'Flora', 'Email': 'Flora.LI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Fountain Lin', 'User First Name': 'Fountain', 'Email': 'Fountain.LIN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Fred Cai', 'User First Name': 'Fred', 'Email': 'Fred.CAI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Freya Hao', 'User First Name': 'Freya', 'Email': 'Freya.Hao@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Gaowei ZHU', 'User First Name': 'Gaowei', 'Email': 'Gaowei.ZHU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Gavin Liu', 'User First Name': 'Gavin', 'Email': 'gavin.liu@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Gilbert Yao (ext.)', 'User First Name': 'Gilbert', 'Email': 'Gilbert.Yao-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Grace Chen', 'User First Name': 'Grace', 'Email': 'grace.chen1@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Grace Lin', 'User First Name': 'Grace', 'Email': 'Grace.LIN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Guangfei ZOU (ext.)', 'User First Name': 'Guangfei', 'Email': 'Guangfei.ZOU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Guozhu Liu', 'User First Name': 'Guozhu', 'Email': 'Guozhu.Liu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Hailong Yang', 'User First Name': 'Hailong', 'Email': 'Hailong.Yang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Haiyan WANG (ext.)', 'User First Name': 'Haiyan', 'Email': 'Haiyan.WANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Hangshi SHENTU (ext.)', 'User First Name': 'Hangshi', 'Email': 'Hangshi.SHENTU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Hanling JIANS (Dummy)', 'User First Name': 'Hanling', 'Email': 'Hanling.JIANS@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Haoyu LIN (ext.)', 'User First Name': 'Haoyu', 'Email': 'Haoyu.LIN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'He Yao (ext.)', 'User First Name': 'He', 'Email': 'He.Yao@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Helen Zheng', 'User First Name': 'Helen', 'Email': 'Helen.ZHENG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Hilde Wong (ext.)', 'User First Name': 'Hilde', 'Email': 'Hilde.Wong@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Holly Chen', 'User First Name': 'Holly', 'Email': 'holly.chen@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Hongbo Yu', 'User First Name': 'Hongbo', 'Email': 'Hongbo.YU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Hongyan ZHAO (ext.)', 'User First Name': 'Hongyan', 'Email': 'HONGYAN.ZHAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Hua CHEN (ext.)', 'User First Name': 'Hua', 'Email': 'Hua.CHEN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Huang Danni', 'User First Name': 'Huang', 'Email': 'Huang.Danni@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Huiqi Mao', 'User First Name': 'Huiqi', 'Email': 'Huiqi.MAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Iris Cheng', 'User First Name': 'Iris', 'Email': 'Iris.CHENG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Iris Zhu', 'User First Name': 'Iris', 'Email': 'Iris.Zhu-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ivy Bi', 'User First Name': 'Ivy', 'Email': 'Ivy.BI@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Izzy Qian', 'User First Name': 'Izzy', 'Email': 'izzy.qian@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jackie Zhou (ext.)', 'User First Name': 'Jackie', 'Email': 'jackie.zhou-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jam Guo', 'User First Name': 'Jam', 'Email': 'Jam.GUO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'James Wang', 'User First Name': 'James', 'Email': 'James.WANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jane Jie', 'User First Name': 'Jane', 'Email': 'Jane.JIE@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Janus Jin', 'User First Name': 'Janus', 'Email': 'Janus.JIN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jason Hu', 'User First Name': 'Jason', 'Email': 'jason.hu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jason Zhao', 'User First Name': 'Jason', 'Email': 'jason.zhao@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jeffery Deng', 'User First Name': 'Jeffery', 'Email': 'Jeffery.DENG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jennifer Cao', 'User First Name': 'Jennifer', 'Email': 'Jennifer.CAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jenny Li (ext.)', 'User First Name': 'Jenny', 'Email': 'jenny.li-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jenny Zuo (Legal) (ext.)', 'User First Name': 'Jenny', 'Email': 'jenny.zuo-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jerry Hong', 'User First Name': 'Jerry', 'Email': 'jerry.hong@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jerry MAO (ext.)', 'User First Name': 'Jerry', 'Email': 'Jerry.MAO@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jesse Wu', 'User First Name': 'Jesse', 'Email': 'Jesse.WU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jessica Wang', 'User First Name': 'Jessica', 'Email': 'jessica.wang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jessica Zhang', 'User First Name': 'Jessica', 'Email': 'jessica.zhang1@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jiaheng WU', 'User First Name': 'Jiaheng', 'Email': 'Jiaheng.WU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jian Chen (ext.)', 'User First Name': 'Jian', 'Email': 'Jian.Chen-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jianguo WANG (ext.)', 'User First Name': 'Jianguo', 'Email': 'Jianguo.Wang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jianming Xu', 'User First Name': 'Jianming', 'Email': 'Jianming.XU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jiaqi Gao (ext.)', 'User First Name': 'Jiaqi', 'Email': 'Jiaqi.GAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jie SUN (ext.)', 'User First Name': 'Jie', 'Email': 'Jie.SUN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jinde Chen', 'User First Name': 'Jinde', 'Email': 'Jinde.CHEN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jing SHIS (dummy)', 'User First Name': 'Jing', 'Email': 'Jing.SHIS@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jingxiong Chen', 'User First Name': 'Jingxiong', 'Email': 'jingxiong.chen@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jinwei HUA (ext.)', 'User First Name': 'Jinwei', 'Email': 'Jinwei.HUA@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Joan Jiang', 'User First Name': 'Joan', 'Email': 'joan.jiang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Joey Gu (ext.)', 'User First Name': 'Joey', 'Email': 'joey.gu-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Joey YAO (ext.)', 'User First Name': 'Joey', 'Email': 'Joey.YAO@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Johnson Zhao', 'User First Name': 'Johnson', 'Email': 'Johnson.ZHAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Jony QUAN (ext.)', 'User First Name': 'Jony', 'Email': 'Jony.QUAN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Juanyu LIAN (ext.)', 'User First Name': 'Juanyu', 'Email': 'Juanyu.LIAN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Julia LIAO', 'User First Name': 'Julia', 'Email': 'Julia.LIAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Julien Proglio', 'User First Name': 'Julien', 'Email': 'Julien.Proglio@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Jun ZHENG', 'User First Name': 'Jun', 'Email': 'Jun.Zheng@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Junmin LIN (ext.)', 'User First Name': 'Junmin', 'Email': 'Junmin.LIN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Kaiming Lin (ext.)', 'User First Name': 'Kaiming', 'Email': 'Kaiming.Lin-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Karen Li', 'User First Name': 'Karen', 'Email': 'karen.li@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Kathie Wang', 'User First Name': 'Kathie', 'Email': 'Kathie.Wang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Kay Cheng', 'User First Name': 'Kay', 'Email': 'Kay.CHENG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Kejia Wu (ext.)', 'User First Name': 'Kejia', 'Email': 'Kejia.WU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Ken Jiang', 'User First Name': 'Ken', 'Email': 'Ken.JIANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Kenneth Tsang', 'User First Name': 'Kenneth', 'Email': 'Kenneth.Tsang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Kevin Lee', 'User First Name': 'Kevin', 'Email': 'A-Chi.Lee@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Kimberley Lu', 'User First Name': 'Kimberley', 'Email': 'kimberley.lu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Kiwi Li (ext.)', 'User First Name': 'Kiwi', 'Email': 'Kiwi.Li-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Kris Liang (ext.)', 'User First Name': 'Kris', 'Email': 'Kris.LIANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Lake Hu (ext.)', 'User First Name': 'Lake', 'Email': 'lake.hu-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Lara Xu (ext.)', 'User First Name': 'Lara', 'Email': 'Lara.XU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Larry PANG', 'User First Name': 'Larry', 'Email': 'Larry.Pang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Leeann Zhou', 'User First Name': 'Leeann', 'Email': 'Leeann.ZHOU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Lei Qin', 'User First Name': 'Lei', 'Email': 'lei.qin1@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Lei TONG', 'User First Name': 'Lei', 'Email': 'Lei.Tong@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Lena Tian', 'User First Name': 'Lena', 'Email': 'lena.tian@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Leo Shi (ext.)', 'User First Name': 'Leo', 'Email': 'leo.shi-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Leon Cui', 'User First Name': 'Leon', 'Email': 'leon.cui@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Leona Wang', 'User First Name': 'Leona', 'Email': 'leona.wang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Li CHEN', 'User First Name': 'Li', 'Email': 'Li.CHEN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Liang Cao', 'User First Name': 'Liang', 'Email': 'Liang.CAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Lijuan Wang', 'User First Name': 'Lijuan', 'Email': 'Lijuan.WANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Lili Jin', 'User First Name': 'Lili', 'Email': 'Lili.JIN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Lily WAN (ext.)', 'User First Name': 'Lily', 'Email': 'Lily.WAN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Lin ZHANG (dummy)', 'User First Name': 'Lin', 'Email': 'Lin.ZHANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Ling Zhang', 'User First Name': 'Ling', 'Email': 'Ling.ZHANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Lise Li', 'User First Name': 'Lise', 'Email': 'lise.li@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Longhui Wang', 'User First Name': 'Longhui', 'Email': 'Longhui.WANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Lorna ZHANG (ext.)', 'User First Name': 'Lorna', 'Email': 'Lorna.ZHANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Lucia Luo (ext.)', 'User First Name': 'Lucia', 'Email': 'Lucia.LUO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Lucy MAI (ext.)', 'User First Name': 'Lucy', 'Email': 'Lucy.MAI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'LuLu Li (ext.)', 'User First Name': 'LuLu', 'Email': 'LuLu.Li@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Lydia Song (ext.)', 'User First Name': 'Lydia', 'Email': 'Lydia.SONG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Madeleine Mao (ext.)', 'User First Name': 'Madeleine', 'Email': 'Madeleine.Mao@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Man Xu', 'User First Name': 'Man', 'Email': 'Man.XU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Maozhou WANG', 'User First Name': 'Maozhou', 'Email': 'Maozhou.Wang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Mark Li', 'User First Name': 'Mark', 'Email': 'Mark.LI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Maro Wang (ext.)', 'User First Name': 'Maro', 'Email': 'Maro.WANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Matthew Liang', 'User First Name': 'Matthew', 'Email': 'Matthew.LIANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Max Zhang (ext.)', 'User First Name': 'Max', 'Email': 'max.zhang-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Mayni Zeng', 'User First Name': 'Mayni', 'Email': 'mayni.zeng@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Meihong FU (ext.)', 'User First Name': 'Meihong', 'Email': 'Meihong.FU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Mia Wang', 'User First Name': 'Mia', 'Email': 'mia.wang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Michael Zhang', 'User First Name': 'Michael', 'Email': 'Michael.Zhang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Michelle Tian', 'User First Name': 'Michelle', 'Email': 'michelle.tian@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Min Wang (ext.)', 'User First Name': 'Min', 'Email': 'min.wang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Mingyan Chen', 'User First Name': 'Mingyan', 'Email': 'Mingyan.CHEN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Monica CHEN (ext.)', 'User First Name': 'Monica', 'Email': 'Monica.CHEN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Mozzie XU (ext.)', 'User First Name': 'Mozzie', 'Email': 'Mozzie.XU-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Nan LI (ext.)', 'User First Name': 'Nan', 'Email': 'Nan.LI2@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Neo Ye', 'User First Name': 'Neo', 'Email': 'Neo.YE@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Nichole Zheng (ext.)', 'User First Name': 'Nichole', 'Email': 'nichole.zheng-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Nick Xie', 'User First Name': 'Nick', 'Email': 'nick.xie@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Nicole Zhou', 'User First Name': 'Nicole', 'Email': 'nicole.zhou1@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Nina Zhu', 'User First Name': 'Nina', 'Email': 'Nina.ZHU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Nino QIAN (ext.)', 'User First Name': 'Nino', 'Email': 'nino.qian-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Paul Doursounian', 'User First Name': 'Paul', 'Email': 'Paul.DOURSOUNIAN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Peng Jiang', 'User First Name': 'Peng', 'Email': 'Peng.Jiang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Peter Li (ext.)', 'User First Name': 'Peter', 'Email': 'Peter.Li-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Philippe Wong', 'User First Name': 'Philippe', 'Email': 'philippe.wong@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ping Wang', 'User First Name': 'Ping', 'Email': 'Ping.WANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'PRC Live', 'User First Name': 'PRC', 'Email': 'prc.live@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Qiang Yu', 'User First Name': 'Qiang', 'Email': 'Qiang.YU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Qing Zhang', 'User First Name': 'Qing', 'Email': 'qing.zhang1@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Qingyong TAO (ext.)', 'User First Name': 'Qingyong', 'Email': 'Qingyong.TAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Queenie Li', 'User First Name': 'Queenie', 'Email': 'Queenie.LI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Rani Ding', 'User First Name': 'Rani', 'Email': 'rani.ding@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ray Xu', 'User First Name': 'Ray', 'Email': 'Ray.XU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Regina Bao', 'User First Name': 'Regina', 'Email': 'Regina.BAO@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Richerd Jin (ext.)', 'User First Name': 'Richerd', 'Email': 'Richerd.Jin-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Rita Liu', 'User First Name': 'Rita', 'Email': 'Rita.LIU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Rock XU (ext.)', 'User First Name': 'Rock', 'Email': 'Rock.XU-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Rolo Luo', 'User First Name': 'Rolo', 'Email': 'Rolo.LUO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Ronghua Su', 'User First Name': 'Ronghua', 'Email': 'Ronghua.Su@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Roy Zhang (Central)', 'User First Name': 'Roy', 'Email': 'Tianyu.ZHANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Sabrina Shen', 'User First Name': 'Sabrina', 'Email': 'Sabrina.SHEN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Sally XU', 'User First Name': 'Sally', 'Email': 'Sally.Xu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Sam Zhao', 'User First Name': 'Sam', 'Email': 'Sam.Zhao@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Sandy Yang', 'User First Name': 'Sandy', 'Email': 'Sandy.Yang@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Sarah Huang', 'User First Name': 'Sarah', 'Email': 'sarah.huang1@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Season Yan', 'User First Name': 'Season', 'Email': 'Season.YAN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Serena Cui', 'User First Name': 'Serena', 'Email': 'Serena.CUI@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Seven Wang', 'User First Name': 'Seven', 'Email': 'Seven.WANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Shane Zhou', 'User First Name': 'Shane', 'Email': 'shane.zhou@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Shawn WANGS (Dummy)', 'User First Name': 'Shawn', 'Email': 'Shawn.WANGS@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Shengbo Qi', 'User First Name': 'Shengbo', 'Email': 'Shengbo.Qi@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'shfesco (ext.)', 'User First Name': 'shfesco', 'Email': 'shfesco@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Shijie Zhang', 'User First Name': 'Shijie', 'Email': 'Shijie.Zhang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Shirley Wang', 'User First Name': 'Shirley', 'Email': 'Shirley.Wong@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Shuai Gao', 'User First Name': 'Shuai', 'Email': 'Shuai.GAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Shuipeng Han', 'User First Name': 'Shuipeng', 'Email': 'Shuipeng.HAN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Silver Shen', 'User First Name': 'Silver', 'Email': 'silver.shen@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Simon Ni', 'User First Name': 'Simon', 'Email': 'Simon.NI@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Sissi LU (ext.)', 'User First Name': 'Sissi', 'Email': 'Sissi.LU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Song GAO (ext.)', 'User First Name': 'Song', 'Email': 'Song.Gao@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Sophie Zhu', 'User First Name': 'Sophie', 'Email': 'sophie.zhu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'SRV.IMPSERVICE3', 'User First Name': 'SRV.IMPSERVICE3', 'Email': 'SRV.IMPSERVICE3@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Star Lin', 'User First Name': 'Star', 'Email': 'Star.LIN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Stella Hu', 'User First Name': 'Stella', 'Email': 'stella.hu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Steven Lang', 'User First Name': 'Steven', 'Email': 'Steven.LANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Steven Zhang (Central)', 'User First Name': 'Steven', 'Email': 'Steven.ZHANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Suirong Chen', 'User First Name': 'Suirong', 'Email': 'suirong.chen@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Summer ZHU (ext.)', 'User First Name': 'Summer', 'Email': 'Summer.ZHU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'System-In-Motion (Dummy)', 'User First Name': 'System-In-Motion', 'Email': 'System.In-Motion@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Tao Xian', 'User First Name': 'Tao', 'Email': 'tao.xian@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Teky Lu', 'User First Name': 'Teky', 'Email': 'Teky.LU@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Teri Ng', 'User First Name': 'Teri', 'Email': 'Teri.NG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Tianjiao Wang (ext.)', 'User First Name': 'Tianjiao', 'Email': 'Tianjiao.Wang-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Tim Liao', 'User First Name': 'Tim', 'Email': 'Tim.LIAO@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Tina Zhang', 'User First Name': 'Tina', 'Email': 'Tina.ZHANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Tony Wu', 'User First Name': 'Tony', 'Email': 'Tony.Wu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Vincent Ge', 'User First Name': 'Vincent', 'Email': 'Vincent.GE@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Vincent Zhang', 'User First Name': 'Vincent', 'Email': 'vincent.zhang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Vivi Zou', 'User First Name': 'Vivi', 'Email': 'vivi.zou@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Vivian Yang', 'User First Name': 'Vivian', 'Email': 'Vivian.YANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Wang Li', 'User First Name': 'Wang', 'Email': 'li.wang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wanxia ZHANG (ext.)', 'User First Name': 'Wanxia', 'Email': 'Wanxia.ZHANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wei Du', 'User First Name': 'Wei', 'Email': 'Wei.Du@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wei Lin', 'User First Name': 'Wei', 'Email': 'wei.lin@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Weidong Wu (ext.)', 'User First Name': 'Weidong', 'Email': 'Weidong.Wu@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wen Fang', 'User First Name': 'Wen', 'Email': 'Wen.FANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wenda DENG', 'User First Name': 'Wenda', 'Email': 'Wenda.Deng@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wendy Li', 'User First Name': 'Wendy', 'Email': 'Wendy.LI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wenjing YU', 'User First Name': 'Wenjing', 'Email': 'Wenjing.YU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wenting XU (ext.)', 'User First Name': 'Wenting', 'Email': 'Wenting.XU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Whiky Jiang', 'User First Name': 'Whiky', 'Email': 'Whiky.JIANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'William Fan (ext.)', 'User First Name': 'William', 'Email': 'William.FAN@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Winni Chen', 'User First Name': 'Winni', 'Email': 'Winni.CHEN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Wood Pan', 'User First Name': 'Wood', 'Email': 'wood.pan@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Xia LI (ext.)', 'User First Name': 'Xia', 'Email': 'Xia.LI@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Xiang LI (ext.)', 'User First Name': 'Xiang', 'Email': 'Xiang.LI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Xiaodong Qian', 'User First Name': 'Xiaodong', 'Email': 'Xiaodong.Qian@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Xiaoman HU (ext.)', 'User First Name': 'Xiaoman', 'Email': 'Xiaoman.HU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Xiaoping WU (ext.)', 'User First Name': 'Xiaoping', 'Email': 'Xiaoping.WU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Xiaoxia WANG (ext.)', 'User First Name': 'Xiaoxia', 'Email': 'Xiaoxia.Wang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Xiaoying LIU', 'User First Name': 'Xiaoying', 'Email': 'Xiaoying.LIU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Xingwang SUN (ext.)', 'User First Name': 'Xingwang', 'Email': 'Xingwang.SUN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Xu Ye (ext.)', 'User First Name': 'Xu', 'Email': 'Xu.Ye-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Xuehong ZHONG (ext.)', 'User First Name': 'Xuehong', 'Email': 'Xuehong.ZHONG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Yahui CHEN (ext.)', 'User First Name': 'Yahui', 'Email': 'Yahui.CHEN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yan Han (ext.)', 'User First Name': 'Yan', 'Email': 'Yan.Han@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yan WANG (ext.)', 'User First Name': 'Yan', 'Email': 'Yan.WANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yang Wang (ext.)', 'User First Name': 'Yang', 'Email': 'Yang.Wang-ext@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Yann Soenen', 'User First Name': 'Yann', 'Email': 'Yann.Soenen@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Yaocong WANG (ext.)', 'User First Name': 'Yaocong', 'Email': 'Yaocong.WANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yeerzhati AZHATI (ext.)', 'User First Name': 'Yeerzhati', 'Email': 'Yeerzhati.AZHATI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yibo Huang', 'User First Name': 'Yibo', 'Email': 'yibo.huang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yinfeng Jin (ext.)', 'User First Name': 'Yinfeng', 'Email': 'Yinfeng.Jin@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Ying Zhang', 'User First Name': 'Ying', 'Email': 'Ying.ZHANG@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Yiqi DAI (ext.)', 'User First Name': 'Yiqi', 'Email': 'Yiqi.DAI@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Yizhu Zhao', 'User First Name': 'Yizhu', 'Email': 'Yizhu.ZHAO@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yong Xia', 'User First Name': 'Yong', 'Email': 'yong.xia@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yonglin ZHANG', 'User First Name': 'Yonglin', 'Email': 'Yonglin.ZHANG@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yongyan CHEN (ext.)', 'User First Name': 'Yongyan', 'Email': 'Yongyan.CHEN@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yu Li', 'User First Name': 'Yu', 'Email': 'Yu.Li-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yuhao Shen', 'User First Name': 'Yuhao', 'Email': 'Yuhao.Shen@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yun Chen1', 'User First Name': 'Yun', 'Email': 'Yun.Chen1@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Yuxin Li', 'User First Name': 'Yuxin', 'Email': 'Yuxin.LI@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Yvonne Liu', 'User First Name': 'Yvonne', 'Email': 'Yvonne.Liu@pernod-ricard.com', 'Shanghai': 'YES'}, {'User': 'Zhang.Gang', 'User First Name': 'Zhang.Gang', 'Email': 'Gang.Zhang-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Zhen Wang', 'User First Name': 'Zhen', 'Email': 'Wang.Zhen@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Zhenjian Wang', 'User First Name': 'Zhenjian', 'Email': 'Zhenjian.Wang@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Zhiquan Chen', 'User First Name': 'Zhiquan', 'Email': 'Zhiquan.Chen@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Zhongxia Wu', 'User First Name': 'Zhongxia', 'Email': 'Zhongxia.WU@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Ziyi Wang (ext.)', 'User First Name': 'Ziyi', 'Email': 'Ziyi.Wang-ext@pernod-ricard.com', 'Shanghai': 'NO'}, {'User': 'Zoya Wang', 'User First Name': 'Zoya', 'Email': 'zoya.wang@pernod-ricard.com', 'Shanghai': 'YES'}]


for i in range(2):
	msg = outlook.GetNamespace("MAPI").OpenSharedItem(r"C:\\Users\\zzhuetan\\OneDrive - PERNOD RICARD\\Desktop\\Check-in 2022 - Basic.msg")
	firstname = userlist[i]['User First Name']
	useremailaddr = userlist[i]['Email']
	msg.HTMLBody = msg.HTMLBody.replace('%%FirstName%%','{},'.format(firstname))

	mail = msg

	mail.To = '{}'.format(useremailaddr)  #收件人
	mail.Subject = '有关IT的支持，可以联系我们！'  #邮件主题
	mail.BodyFormat = 2  # 2表示使用Html format，可以调整格式

	mail.Display()  #显示发送邮件界面