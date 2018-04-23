print("""
####################
#文件时间戳生成工具#
####################
#   作者：hyneman  #
####################
""")
import win32ui
import time
import sys  
from Crypto.Cipher import AES  
from binascii import b2a_hex, a2b_hex  
import hashlib
import zmail
def gettimestr():#取得时间戳
    t = time.time()
    t = time.localtime(t)
    tl = list(t)
    t_str = str(tl[0])+"-"+str(tl[1])+"-"+str(tl[2])+"-"+str(tl[3])+"-"+str(tl[4])+"-"+str(tl[5]) 
    return t_str

def fileaddress():#弹出打开文件对话框
    try:
        dlg = win32ui.CreateFileDialog(1, '','test.xlsx') # 1表示打开文件对话框
        dlg.DoModal()
        filename = dlg.GetPathName()
    except Exception as e:
        print(str(e))
        filename = fileaddress()
    return filename

def hash_md5(file_ad):
    BLOCKSIZE = 65536
    hasher = hashlib.md5()
    with open(file_ad,"rb") as a:
        buf = a.read(BLOCKSIZE)
        while len(buf) > 0:
            hasher.update(buf)
            buf = a.read(BLOCKSIZE)
    return hasher.hexdigest()


def hash_sha256(file_ad):
    BLOCKSIZE = 65536
    hasher = hashlib.sha256()
    with open(file_ad,"rb") as a:
        buf = a.read(BLOCKSIZE)
        while len(buf) > 0:
            hasher.update(buf)
            buf = a.read(BLOCKSIZE)
    return hasher.hexdigest()

class prpcrypt():  
    def __init__(self, key):  
        self.key = key  
        self.mode = AES.MODE_CBC  
       
    #加密函数，如果text不是16的倍数【加密文本text必须为16的倍数！】，那就补足为16的倍数  
    def encrypt(self, text):  
        cryptor = AES.new(self.key, self.mode, self.key)  
        text = text.encode("utf-8")  
        #这里密钥key 长度必须为16（AES-128）、24（AES-192）、或32（AES-256）Bytes 长度.目前AES-128足够用  
        length = 16  
        count = len(text)  
        add = length - (count % length)  
        text = text + (b'\0' * add)  
        self.ciphertext = cryptor.encrypt(text)  
        #因为AES加密时候得到的字符串不一定是ascii字符集的，输出到终端或者保存时候可能存在问题  
        #所以这里统一把加密后的字符串转化为16进制字符串  
        return b2a_hex(self.ciphertext).decode("ASCII")  
       
    #解密后，去掉补足的空格用strip() 去掉  
    def decrypt(self, text):  
        cryptor = AES.new(self.key, self.mode, self.key)  
        plain_text = cryptor.decrypt(a2b_hex(text))  
        return plain_text.rstrip(b'\0').decode("utf-8")  
def set_ini():
    emadr = input("e-mail:")
    passwd = input("password:")
    data = "{}--{}".format(emadr,passwd)
    pc = prpcrypt('asdfghjklqwertyu')
    e = pc.encrypt(data)
    with open("UserConfig.ini","w") as f:
            f.write(e)
    return emadr,passwd

def initialize():
    try:
        pc = prpcrypt('asdfghjklqwertyu')
        with open("UserConfig.ini","r") as f:
            data = pc.decrypt(f.read())
            user_inf = data.split("--")
    except Exception :
        emadr,passwd = set_ini()
        user_inf = [emadr,passwd]
    return user_inf

user_inf= initialize()
tstr=gettimestr()
file_ad=fileaddress()
name = input("邮件名：")
print("开始计算文件哈希值...")
sha_str = hash_sha256(file_ad)
md5_str = hash_md5(file_ad)
print("计算完成，正在发生邮件")
text = "file:"+file_ad +"\ntime:" + tstr + "\nSHA256:" + sha_str + "\nMD5:"+ md5_str
mail = {
    'subject': name,  # Anything you want.
    'content': text,  # Anything you want.
}
server = zmail.server(user_inf[0],user_inf[1])
server.send_mail(user_inf[0], mail)
input("时间戳生成完成，按回车退出")
