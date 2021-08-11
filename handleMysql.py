import datetime
import json
import yagmail
from sshtunnel import SSHTunnelForwarder
import pymysql
from handleExcel import handleExcel
from dataTimeEncoder import DateTimeEncoder


class writeLogToExcel:
    def __init__(self):
        today = datetime.date.today()
        oneday = datetime.timedelta(days=1)
        self.yesterday = today - oneday
        self.excelName = '{}莫斯系统日志及统计.xls'.format(self.yesterday)
        self.excel = handleExcel(self.excelName)

    def handleMysql(self):
        server = SSHTunnelForwarder(
            # 跳板机IP，跳板机端口号
            ('47.102.144.10', 22),
            ssh_username='root',
            ssh_password='R28OV7qdhPVQiLMI',
            remote_bind_address=('rm-uf6644gsugn59l8qs.mysql.rds.aliyuncs.com', 3306)
        )
        server.start()
        # host必须为127.0.0.1，代表本机(堡垒机)，user和passwd填的是远程数据库的账号密码
        conn = pymysql.connect(host='127.0.0.1', port=server.local_bind_port, user='mars_view',
                               passwd='Bwxg3yf5Qj9kmXCb',
                               db='mars_report', charset='utf8')
        # 创建游标
        cur = conn.cursor()
        # 执行sql语句
        cur.execute(
            query="select name'用户名',username'用户手机号',operation'动作',gmt_created'操作时间',case type when 1 then'android' when 2 then 'ios' when 3 then 'h5' when 0 then 'pc' else '未登录' end '终端类型'from sys_log where sys_log.log_id > (select log_id from sys_log where gmt_created < curdate() + interval -1 day order by log_id desc limit 1)and gmt_created < curdate()order by log_id desc")
        # 返回所有结果
        res = cur.fetchall()
        sheet1 = json.dumps(res, cls=DateTimeEncoder, ensure_ascii=False)
        cur.execute(
            "select sys_user.name'用户名',sys_user.user_name'手机号',count(1)'总访问次数',count(if(sys_log.type = 0, 1, null))'web',count(if(sys_log.type = 1, 1, null))'android',count(if(sys_log.type = 2, 1, null))'ios',count(if(sys_log.type = 3, 1, null))'h5',remark,(select group_concat(sc.name)from sys_company sc join sys_user_company on sys_user_company.company_id = sc.id where sys_user_company.user_id = sys_user.user_id) '所有企业',ifnull(sys_company.name, '个人')'当前企业'from sys_log join sys_user on sys_user.user_id = sys_log.user_id left join sys_company on sys_company.id = sys_user.company_id where sys_log.log_id > (select log_id from sys_log where gmt_created < curdate()+ interval -1 day order by log_id desc limit 1 ) and sys_log.gmt_created < curdate()group by sys_log.user_id")
        # 返回所有结果
        res = cur.fetchall()
        sheet2 = json.dumps(res, cls=DateTimeEncoder, ensure_ascii=False)
        # 关闭游标
        cur.close()
        # 关闭连接
        conn.close()
        # 关闭服务
        server.close()
        return sheet1, sheet2

    def writeToSheet1(self):
        try:
            data = json.loads(self.handleMysql()[0])
            self.excel.writeSheet1(self.excelName, data)
        except Exception as e:
            print(e)
        finally:
            print("sheet1写入成功")

    def writeToSheet2(self):
        try:
            data = json.loads(self.handleMysql()[1])
            self.excel.writeSheet2(self.excelName, data)
        except Exception as e:
            print(e)
        finally:
            print("sheet2写入成功")

    def sortExcel(self):
        try:
            yagmail.SMTP(
                host='smtp.qq.com', user='1140944259@qq.com',
                password='hhgyerdbjvemiefg', smtp_ssl=True
            ).send('wfy@mars-tech.com.cn', '{}莫斯简报用户操作日志及统计'.format(self.yesterday), self.excelName)
        except Exception as e:
            print(e)
        finally:
            print('邮件发送成功')


if __name__ == '__main__':
    w = writeLogToExcel()
    w.writeToSheet1()
    w.writeToSheet2()
    w.sortExcel()
