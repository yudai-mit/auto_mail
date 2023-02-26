#モジュールのimport、soundは音を鳴らすだけ
import win32com.client
import pandas as pd
import schedule
import time
from sound import sound1


otl=win32com.client.Dispatch("Outlook.Application")
outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace('MAPI')
#受信トレイの指定
inbox=outlook.GetDefaultFolder(6)
#メールの数を比較するためのリスト
lst=[]
#メールを自動送信する関数
def mail_send(mail_Subject):
    
    #メールの内容
    mail=otl.CreateItem(0)

    mail.to='zeromahou@icloud.com'
    mail.subject=str(mail_Subject)+'について自動返信です。'
    mail.Cc='zeromahou@outlook.jp'
    mail.bodyFormat=1
    mail.body='''お忙しい中ご連絡ありがとうございます。
    こちら自動返信となっております。受信は正常に行われました。
    内容の確認、準備ができ次第返信させていただきます。'''
    #送信する前の確認の際はコメントを外す
    #mail.display(True)
    mail.Send()
    print('送信しました')


#メールを受信する関数
def mail_receive(lst):
    #比較方法の都合で0をlstに入れる
    if len(lst)<3:
        for i in range(0,3):
            lst.append(0)
            
    print('Name:',inbox.name)
    print('Count:'+str(len(inbox.Items)))
    
    messages=inbox.Items
    #メールをデータフレームに格納
    df_mail=pd.DataFrame()
    i=0
    for message in messages:
        df_mail.loc[i,"receivedtime"]=str(message.ReceivedTime)[:-6]
        df_mail.loc[i,"sender"]=str(message.SenderEmailAddress)
        df_mail.loc[i,"subject"]=str(message.Subject)
        df_mail.loc[i,"body"]=str(message.body)
        i+=1

    #メールの情報を書き込むcsvファイルを指定
    file_name='mail.csv'
    file_name_all='mail_all.csv'
    
    #自動送信したい人のメールアドレスを指定、情報を抽出
    df_want=df_mail.query('sender=="zeromahou@icloud.com"')
  
    #csvファイルに書き込み
    df_want.to_csv(file_name,encoding="shift-jis",errors='ignore')
    df_mail.to_csv(file_name_all,encodeing='shift-jis',errors='ignore')
    
    #メールが新規に届いたか確認に必要な情報
    mail_num=len(df_want)
    lst.append(mail_num)
    
    #新規に届いていた場合の処理
    if lst[len(lst)-2]<lst[len(lst)-1]:
        #一番新しく届いたメールの件名をデータフレームから抽出しmail_send関数に渡す
        mail_Object=df_want['subject'].tail(1)
        mail_Subject=mail_Object.iloc[-1]
        mail_send(mail_Subject)
    #メール確認したことを知らせる通知音
    sound1()   
#毎時0分にメールが新着していないか確認する理想は毎秒だが私のPCのスペック上処理が重たいので
schedule.every().hours.at(":00").do(mail_receive,lst)


    
#テスト用
"""
for i in range(0,3):   
    mail_receive(lst) 

schedule.every().minutes.at(":00").do(mail_receive,lst)    
"""
#コンピューターの動作間隔
while True:
    schedule.run_pending()
    time.sleep(1)
