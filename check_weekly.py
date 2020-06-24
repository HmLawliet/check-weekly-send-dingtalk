'''
每周日晚十点提醒未发周报的同事发送周报
每周一早九点提醒管理者统计未发周报的同事
'''
import poplib
from email.parser import BytesParser
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
from email.message import EmailMessage
import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_ERROR
import logging
import time

from dingtalk.client import  AppKeyClient
from dingtalk.client.api import Department,User
from dingtalkchatbot.chatbot import DingtalkChatbot



# 周报组的邮箱信息
WEEKLY_GROUP = 'report_hz@wopuwulian.com'
# 钉钉的公司唯一id
CORPID = 'dingbcf5486299627a689'  
# 后台应用的appkey
APPKEY = 'ding2aql6y15ldrjo9rzs'  
# 后台应用的appsecret
APPSECRET = 'tc5VNrT4dKQUIv0ZjVQjQ9Ma8Q86gEuOkJnZonbQaldGvWk7Cx4dnK9vFFtNMB_-Is' 
# 软件部门id
DEPARTMENT_ID = '1104410500'
# 统计周报的负责人id
GROUP_CONSTRUCTION_MANAGER_mobile = '18070142342'
# 输入邮件地址, 口令和POP3服务器地址:
USERNAME_ = 'houmin@wopuwulian.com'
PASSWD_ = '******'
SERVER_ = 'pop3.wopuwulian.com'

# 钉钉连接
WEBHOOK = 'https://oapi.dingtalk.com/robot/send?access_token=t7fd1f0342913bde4db25d34b57692a7202de68dca1c8682b328c36c7ce6b25a5s'
SECRET = 'TSECac249cbaa61d533e6599e2b289cb7acc37dc5686babe924803e4b9ca05a63187S'

def week_range(weekdelta=-1):
    '''
    周的起止时间
    :param weekdelta 偏移量    0本周, -1上周, 1下周
    :return tuple : （周一,周日）
    '''
    now = datetime.datetime.now()
    week = now.weekday()
    from_ = (now - datetime.timedelta(days=week - 7 * weekdelta)).date()
    to_ = (now + datetime.timedelta(days=6 - week + 7 * weekdelta)).date()
    return from_, to_

def decode_str(s):
    '''
    邮件解码字符串
    :param s ： 要解析的字符串
    :return  解析后的值
    '''
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


def parseweekmail(el,pl,st):
    '''
    :param el 邮箱长度
    :param pl poplib server对象
    :param st 解析周报的开始时间
    :return 邮箱列表
    '''
    sender_list = []
    for index in range(el,0,-1):
        lines = pl.retr(index)[1]
        msg = BytesParser(EmailMessage).parsebytes(b'\r\n'.join(lines))
        
        # 判断是否是本周  判断是否接受者是周报组
        mail_date = parsedate_to_datetime(msg.get('Date', "")).date()
        mail_receiver = parseaddr(msg.get('To', ""))[1]
        mail_cc = parseaddr(msg.get('Cc',""))[1]
        if mail_date < st:
            break 
        mail_subject = decode_str(msg.get('Subject', "")) 
        if (mail_receiver == WEEKLY_GROUP or WEEKLY_GROUP in mail_cc ) and not (
            mail_subject.startswith('项目周报') or 
            decode_str(mail_subject).split('(')[0].endswith('项目周报') or 
            decode_str(mail_subject).split('（')[0].endswith('项目周报')
        ):
            sender_list.append(parseaddr(msg.get('From', ""))[1]) 
    return sender_list
        

def getMail():
    '''获取邮件'''
    today = datetime.datetime.now()
    if today.weekday() == 6: # 周日提醒ds=at_list)
        st,_ = week_range(0)
    elif today.weekday() == 0:  # 周一惩罚
        st,_ = week_range()
    else:
        # 测试用  正式 应该 st = None
        st = None # 测试 st,_ = week_range()  
    if not st:
        return 
    # 连接到POP3服务器:
    pl = poplib.POP3(SERVER_)
    # 关闭调试信息:
    pl.set_debuglevel(0)
    # 身份认证:
    pl.user(USERNAME_)
    pl.pass_(PASSWD_)
    # 邮件长度
    el = len(pl.list()[1])
    sd = parseweekmail(el,pl,st)
    # 关闭连接:
    pl.quit()
    return sd 


def all_department(root_id,dp_obj,dp_lst):
    '''
    递归遍历root部门下的所有子部门目录下的部门id
    '''
    temp_list = dp_obj.list_ids(root_id)
    for item in temp_list:
        if item == root_id:
            continue
        res = all_department(item,dp_obj,dp_lst)
        if not res:
            dp_lst.append(item)
            continue
        return dp_lst
    
def software_sender():
    '''
    从钉钉中获取软件部门的邮箱与手机信息
    '''
    # 获取钉钉那边的软件部门用户信息  
    # 暂时没有权限
    client = AppKeyClient(CORPID,APPKEY,APPSECRET)
    # 查询所有软件部门下的所有子目录包含父目录
    dp_list = []
    dp = Department(client)
    all_department(DEPARTMENT_ID,dp,dp_list)
    dp_list.append(DEPARTMENT_ID)
    # 查询所有的用户id，以及详情信息
    if not dp_list:
        return 
    us_dict = {}
    us = User(client)
    for dp_id in dp_list:
        r = us.list(dp_id)
        if r['errmsg'] == 'ok':
            for i in r['userlist']:
                try:
                    # orgEmail 公司邮箱  userid 用户id 有些id不全  所以改用手机号 mobile 手机号
                    if not i.get('orgEmail','') in us_dict.keys() and i['position'] !='总监':
                        us_dict[i['orgEmail']] = i['mobile']
                except Exception as e:
                    print(i['name'],e)
                    continue
    return us_dict

def send_at_msg(at_list,at_all=False):
    '''
    发送钉钉消息并@群内成员
    :param at_list @成员的手机号
    :param at_all @所有人  默认为False  
    :return None
    '''
    dt = DingtalkChatbot(webhook=WEBHOOK,secret=SECRET)
    today = datetime.datetime.now()
    if today.weekday() == 6: # 周日提醒
        msg = f'截至「 {today.strftime("%Y-%m-%d HH:mm:ss")} 」还未发报告未发周报，请及时发送至周报组！'
        dt.send_text(msg=msg,at_mobiles=at_list)
    elif today.weekday() == 0:  # 周一惩罚
        if at_all:
            dt.send_text(msg='周报上缴齐全!',is_at_all=at_all)
            return 
        dt.send_text(msg='慷慨解囊积极分子，团建制度的忠实拥户者！',at_mobiles=at_list)
        dt.send_text(msg='请财大大速速处理！',at_mobiles=[GROUP_CONSTRUCTION_MANAGER_mobile,])
    else:
        pass
        # # 测试用
        # dt.send_text(msg='慷慨解囊积极分子，团建制度的忠实拥户者！',at_mobiles=at_list)
        # dt.send_text(msg='请财大大速速处理！',at_mobiles=[GROUP_CONSTRUCTION_MANAGER_mobile,])
         
    
def main():
    sd_list = getMail()
    if not sd_list:
        return 
    ssd = software_sender()
    soft_sender = []
    for key,_ in ssd.items():
        if key in sd_list:
            soft_sender.append(key)
    no_sender = list(set(list(ssd.keys())) - set(soft_sender))
    # 全部发送了周报就不需要钉钉提醒
    if not no_sender:
        return 
    # 获取到未发送人的手机号进行@提示
    notice_sender = [ ssd.get(item) for item in no_sender ]     
    if not notice_sender:
        return 
    send_at_msg(notice_sender)
    print('发送提醒成功！！！')
    

class MonitorModel(object):
    def __init__(self, level=logging.INFO):
        self.scheduler = BlockingScheduler()
        self.logger = MonitorModel.initlogger(level=level)

    def listerner(self, event):
        '''
        当job抛出异常时，APScheduler会默默的把他吞掉，不提供任何提示，这不是一种好的实践，我们必须知晓程序的任何差错。
        APScheduler提供注册listener，可以监听一些事件，包括：job抛出异常、job没有来得及执行等。
        '''
        if event.exception:
            self.logger.error('任务出错了！！！！！！')
            self.logger.error('暂停')
            self.scheduler.pause()
            self.logger.error('重启进程')
            self.logger.error('继续')
            self.scheduler.resume()
        else:
            self.logger.info('任务照常运行...')

    def run(self):
        
        # 周日晚10点提醒
        condition_cron = CronTrigger(day_of_week=6,hour=22,minute=0)
        self.scheduler.add_job(main, condition_cron)

        # 周一早9点惩罚
        subscribe_cron = CronTrigger(day_of_week=0,hour=9,minute=30)
        self.scheduler.add_job(main, subscribe_cron)

        # 监听
        self.scheduler.add_listener(self.listerner, EVENT_JOB_EXECUTED | EVENT_JOB_ERROR)
        self.scheduler._logger = logging
        
        self.scheduler.start()

    @staticmethod
    def initlogger(level):
        '''
        初始化日志配置
        '''
        # 第一步，创建一个logger
        logging.basicConfig()
        logger = logging.getLogger()
        logger.setLevel(level)  # Log等级总开关
        # 第二步，创建一个handler，用于写入日志文件
        rq = time.strftime('%Y%m%d%H%M', time.localtime(time.time())) + 'SchedulerLogs'
        # 日志目录
        logfile = f'{rq}.log'
        fh = logging.FileHandler(logfile, mode='w')
        fh.setLevel(level)  # 输出到file的log等级的开关
        # 第三步，定义handler的输出格式
        formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
        fh.setFormatter(formatter)
        # 第四步，将logger添加到handler里面
        logger.addHandler(fh)
        return logger


if __name__ == "__main__":
    # 定时监控
    monitor = MonitorModel(level=logging.INFO)
    monitor.run()
    
    # main()
    
    
