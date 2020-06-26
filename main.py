#6/17~
#-------全体フロー作成-------------

#-*- coding :utf-8 -*-
from __future__ import print_function
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import (CarouselColumn, CarouselTemplate, FollowEvent,ConfirmTemplate,
    LocationMessage, MessageEvent, TemplateSendMessage,PostbackEvent,PostbackTemplateAction,
    TextMessage, TextSendMessage, UnfollowEvent, URITemplateAction, ButtonsTemplate, DatetimePickerTemplateAction)
import os
import json
import urllib.parse
import urllib.request
import requests
import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import datetime




app=Flask(__name__)
line_bot_api=LineBotApi(os.getenv("LINE_CHANNEL_ACCESS_TOKEN"))
handler=WebhookHandler(os.getenv("LINE_CHANNEL_SECRET"))

FOLLOWED_MESSAGE="フォローありがとうございます！\nこちらのアカウントから予約をすることができますので当店ご利用の際はぜひお試しください"+chr(0x10005C)


@app.route('/callback',methods=["POST"])
def callback():
    signature=request.headers['X-Line-Signature']
    body=request.get_data(as_text=True)
    try:
        handler.handle(body,signature)
    except InvalidSignatureError:
        about(400)
    except LineBotApiError as e:
        app.logger.exception(f'LineBotApiError: {e.status_code} {e.message}', e)
        raise e

    return 'OK'



@handler.add(MessageEvent, message=TextMessage)
def handle_text_message(event):
    if event.message.text=="予約したい":
        profile = line_bot_api.get_profile(event.source.user_id)
        userId=profile.user_id
        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="メッセージありがとうございます、ご予約に関してですね"+chr(0x10008D)
           +"\n以下からまずメニューをお選びください！")
        )
        #line_bot_api.reply_message(event.reply_token,TextSendMessage(text="次からスタイリストを選択してください"))
        response_json_list=[]
        result_dict={
            "thumbnail_image_url": "https://imgbp.hotp.jp/CSP/IMG_SRC/55/12/B060655512/B060655512_271-361.jpg",
            "title": "カット",
            "text": "カットにシャンプー・ブロー込のコースです！",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu1"
                }}
        response_json_list.append(result_dict)
        result_dict2={
            "thumbnail_image_url": "https://cloudinary-a.akamaihd.net/vivivi/image/upload/t_beauty,f_auto,dpr_2.0,q_auto:good/c_pad,w_500,h_500/v1574838084/stl_43760_5dde1f433079f_01.jpg",
            "title": "カラーリング",
            "text": "トリートメントつきカラーリングコースです！長さ一律料金です",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu2"
                }}
        response_json_list.append(result_dict2)
        result_dict3={
            "thumbnail_image_url": "https://d3kszy5ca3yqvh.cloudfront.net/wp-content/uploads/2019/3/29/16/083d03f7740809c632d20f5e9d6afc86.jpg",
            "title": "有機デジタルパーマ",
            "text": "持ちが良く髪に負担をかけにくい有機デジタルパーマのコースです！",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu3"
                }}
        response_json_list.append(result_dict3)
        result_dict4={
            "thumbnail_image_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcRtkxvQ8nQ9-DgHCmcgPw23g_4mJzjxOwvH49c9IIbrqLP-YV0J&usqp=CAU",
            "title": "ヘッドスパ",
            "text": "自分へのご褒美に！ヘッドスパ、有機炭酸シャンプー、トリートメントのコース",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu4"
                }}
        response_json_list.append(result_dict4)

        columns = [
            CarouselColumn(
                thumbnail_image_url=column["thumbnail_image_url"],
                title=column["title"],
                text=column["text"],
                actions=[
                PostbackTemplateAction(
                    label=column["actions"]["label"],
                    data=column["actions"]["data"]
                ),
            ])
            for column in response_json_list
        ]
        messages = TemplateSendMessage(
            alt_text="メニュー選択",
            template=CarouselTemplate(columns=columns),
        )
        print("messages is: {}".format(messages))

        line_bot_api.reply_message(
            event.reply_token,
            messages=messages
        )

    elif "・お名前" in event.message.text:
        finalName=event.message.text.lstrip("・お名前: ")[:-18]
        finalNumber=event.message.text[-10:]
        profile = line_bot_api.get_profile(event.source.user_id)
        userId=profile.user_id
        scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('sheets.json', scope)
        gc = gspread.authorize(credentials)
        wks = gc.open('美容院予約 bot sample').sheet1
        line_number=[]
        for i in range(2,len(wks.col_values(1))+2):
            if userId==wks.acell('A'+str(i)).value:
                wks.update_acell('I'+str(i),"selected")
                line_number.append(i)
                break
            elif wks.acell('A'+str(i)).value=="":
                wks.update_acell('A'+str(i),userId)
                wks.update_acell('I'+str(i),"selected")
                line_number.append(i)
                break
        wks.update_acell('F'+str(line_number[0]),finalName)
        wks.update_acell('G'+str(line_number[0]),finalNumber)
        postback_date=wks.acell('D'+str(line_number[0])).value
        postback_time=wks.acell('E'+str(line_number[0])).value
        postback_menu=wks.acell('B'+str(line_number[0])).value

        if postback_menu=="menu1":
            finalMenu="カット"
        elif postback_menu=="menu2":
            finalMenu="カラーリング"
        elif postback_menu=="menu3":
            finalMenu="有機デジタルパーマ"
        elif postback_menu=="menu4":
            finalMenu="ヘッドスパ"
        #----------googleカレンダー転記
        printCalendar(userId,finalName,postback_date,postback_time,finalMenu)

        reserve_confirmation=ButtonsTemplate(
           #title='予約最終確認',
           text='以下の内容で予約を確定します。\n・お名前: '+finalName+"・日時: "+postback_date+" "+postback_time+"\n・メニュー: "
                    +finalMenu,
           actions=[
               PostbackTemplateAction(
                   label='予約完了する',
                   data='reserveComplete'
               ),
               PostbackTemplateAction(
                   label='予約変更を行う',
                   data='changeReserve'
               )
           ]
        )
        line_bot_api.push_message(
           to=userId,
           messages=TemplateSendMessage(
               alt_text='button template',
               template=reserve_confirmation
           )
        )
    elif event.message.text=="予約を変更したい":
        profile = line_bot_api.get_profile(event.source.user_id)
        userId=profile.user_id
        scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('sheets.json', scope)
        gc = gspread.authorize(credentials)
        wks = gc.open('美容院予約 bot sample').sheet1
        line_number=[]
        for i in range(2,len(wks.col_values(1))+2):
            if userId==wks.acell('A'+str(i)).value:
                wks.update_acell('I'+str(i),"selected")
                line_number.append(i)
                break
            elif wks.acell('A'+str(i)).value=="":
                wks.update_acell('A'+str(i),userId)
                wks.update_acell('I'+str(i),"selected")
                line_number.append(i)
                break
        finalName=wks.acell('F'+str(line_number[0])).value
        postback_date=wks.acell('D'+str(line_number[0])).value
        postback_time=wks.acell('E'+str(line_number[0])).value
        postback_menu=wks.acell('B'+str(line_number[0])).value
        if postback_menu=="menu1":
            finalMenu="カット"
        elif postback_menu=="menu2":
            finalMenu="カラーリング"
        elif postback_menu=="menu3":
            finalMenu="有機デジタルパーマ"
        elif postback_menu=="menu4":
            finalMenu="ヘッドスパ"
        else:
            finalMenu="データなし"
        changeReservation=ButtonsTemplate(
           #title='予約最終確認',
           text='現在の予約内容は\n・お名前: '+finalName+"・日時: "+postback_date+" "+postback_time+"\n・メニュー: "
                    +finalMenu +"です。",
           actions=[
               PostbackTemplateAction(
                   label='変更しない',
                   data='reserveComplete'
               ),
               PostbackTemplateAction(
                   label='日にちを変更する',
                   data='changeDate'
               ),
               PostbackTemplateAction(
                   label='時間を変更する',
                   data='reserveTime2'
               ),
               PostbackTemplateAction(
                   label='メニューを変更する',
                   data='reserveComplete'
               ),
           ]
        )
        line_bot_api.push_message(
           to=userId,
           messages=TemplateSendMessage(
               alt_text='button template',
               template=changeReservation
           )
        )



    else:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=event.message.text)
        )


#---------postback------------------
@handler.add(PostbackEvent)
def getResponse(event):
    profile = line_bot_api.get_profile(event.source.user_id)
    userId=profile.user_id
    postback_msg=event.postback.data
    scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('sheets.json', scope)
    gc = gspread.authorize(credentials)
    wks = gc.open('美容院予約 bot sample').sheet1
    line_number=[]
    for i in range(2,len(wks.col_values(1))+2):
        if userId==wks.acell('A'+str(i)).value:
            wks.update_acell('I'+str(i),"selected")
            line_number.append(i)
            break
        elif wks.acell('A'+str(i)).value=="":
            wks.update_acell('A'+str(i),userId)
            wks.update_acell('I'+str(i),"selected")
            line_number.append(i)
            break
    if postback_msg=='menu1' or postback_msg=='menu2' or postback_msg=='menu3' or postback_msg=='menu4':
        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="次にスタイリストを選択してください"+chr(0x10008D))
        )
        wks.update_acell('B'+str(line_number[0]),postback_msg)
        response_json_list=[]
        result_dict={
            "thumbnail_image_url": "https://www.marunage.co.jp/media/wp-content/uploads/2016/08/1606_biyoushiwatanabe_01.jpg",
            "title": "木村弘",
            "text": "スタイリスト歴８年",
            "actions": {
                "label": "この美容師を指定する",
                "data":"stylist1"
                }}
        response_json_list.append(result_dict)
        result_dict2={
            "thumbnail_image_url": "https://www.co-mirrorball.com/wp-content/uploads/2018/06/20cfce153c6dcbd30dc35695758066ae-1.jpg",
            "title": "山田花子",
            "text": "スタイリスト歴3年",
            "actions": {
                "label": "この美容師を指定する",
                "data":"stylist2"
                }}
        response_json_list.append(result_dict2)
        result_dict3={
            "thumbnail_image_url": "https://myogadani.zaza-j.com/upload/tenant_1/7eefd4685390fdd055a75ad14015aa37.jpg",                "title": "指定しない",
            "text": "指定なしにすると対応可能な時間が増えます",
            "actions": {
                "label": "指定しないで進む",
                "data":"stylist_none"
                }}
        response_json_list.append(result_dict3)

        columns = [
            CarouselColumn(
                thumbnail_image_url=column["thumbnail_image_url"],
                title=column["title"],
                text=column["text"],
                actions=[
                PostbackTemplateAction(
                    label=column["actions"]["label"],
                    data=column["actions"]["data"]
                ),
            ])
            for column in response_json_list
        ]
        messages = TemplateSendMessage(
            alt_text="次からスタイリストを選択してください",
            template=CarouselTemplate(columns=columns),
        )
        print("messages is: {}".format(messages))
        line_bot_api.push_message(
           to=userId,
           messages=messages
        )

    elif postback_msg=='stylist1':
        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="スタイリスト木村弘の予約状況を調べます、しばらくお待ちください"+chr(0x100079))
        )
        wks.update_acell('C'+str(line_number[0]),postback_msg)
        holiday=getHoliday(1)
        holidayAll=""
        for i in range(len(holiday)):
            if i==len(holiday)-1:
                _holiday=holiday[i].replace("-","/")
                holidayAll+=_holiday
            else:
                _holiday=holiday[i].replace("-","/")
                holidayAll+=_holiday+","

        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="木村は"+holidayAll+"以外出勤しています！\n以下から希望日を選択ください"+chr(0x10008D))
        )
        today=datetime.date.today().strftime("%Y-%m-%d")
        date_picker = TemplateSendMessage(
            alt_text='予定日を設定',
            template=ButtonsTemplate(
                text='以下から選択して下さい',
                title='希望日を選択してください',
                actions=[
                        DatetimePickerTemplateAction(
                        label='ここをタップ',
                        data='reserveTime',
                        mode='date',
                        initial=today,
                        min=today,
                        max='2099-12-31'
                    )
                ]
            )
        )
        line_bot_api.push_message(
           to=userId,
           messages=date_picker)

    elif postback_msg=='stylist2':
        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="スタイリスト山田の予約状況を調べます、しばらくお待ちください"+chr(0x100079))
        )
        wks.update_acell('C'+str(line_number[0]),postback_msg)
        holiday=getHoliday(2)
        holidayAll=""
        for i in range(len(holiday)):
            if i==len(holiday)-1:
                _holiday=holiday[i].replace("-","/")
                holidayAll+=_holiday
            else:
                _holiday=holiday[i].replace("-","/")
                holidayAll+=_holiday+","

        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="木村は"+holidayAll+"以外出勤しています！\n以下から希望日を選択ください"+chr(0x10008D))
        )
        today=datetime.date.today().strftime("%Y-%m-%d")
        date_picker = TemplateSendMessage(
            alt_text='予定日を設定',
            template=ButtonsTemplate(
                text='以下から選択して下さい',
                title='希望日を選択してください',
                actions=[
                        DatetimePickerTemplateAction(
                        label='ここをタップ',
                        data='reserveTime',
                        mode='date',
                        initial=today,
                        min=today,
                        max='2099-12-31'
                    )
                ]
            )
        )
        line_bot_api.push_message(
           to=userId,
           messages=date_picker)


    # elif postback_msg=='reservation1':
    #      postback_date=event.postback.params['date']
    #      wks.update_acell('D'+str(line_number[0]),postback_date)
    #      date_confirmation=ButtonsTemplate(
    #         title='日にち入力確認',
    #         text=postback_date+'でお間違い無いでしょうか。',
    #         actions=[
    #             PostbackTemplateAction(
    #                 label='日にちを変更する',
    #                 data='changeDate'
    #             ),
    #             PostbackTemplateAction(
    #                 label='次に時間を選択する',
    #                 data='reserveTime'
    #             ),
    #
    #         ]
    #      )
    #      line_bot_api.push_message(
    #         to=userId,
    #         messages=TemplateSendMessage(
    #             alt_text='button template',
    #             template=date_confirmation
    #
    #         )
    #
    #      )
    elif  postback_msg=='changeDate':
         today=datetime.date.today().strftime("%Y-%m-%d")
         date_picker = TemplateSendMessage(
             alt_text='予定日を設定',
             template=ButtonsTemplate(
                 text='以下から選択して下さい',
                 title='希望日を選択して下さい。',
                 actions=[
                     DatetimePickerTemplateAction(
                         label='ここをタップ',
                         data='reservation3',
                         mode='date',
                         initial=today,
                         min=today,
                         max='2099-12-31'
                     )
                 ]
             )
         )
         line_bot_api.push_message(
            to=userId,
            messages=date_picker
         )

#reserve time
    elif postback_msg=='reserveTime':
         postback_date=event.postback.params['date']
         wks.update_acell('D'+str(line_number[0]),postback_date)
         reservedStylist=wks.acell('C'+str(line_number[0])).value
         reservedTime=getTime(reservedStylist,postback_date)
         if reservedTime=="None":
             line_bot_api.push_message(
                to=userId,
                messages=TextSendMessage(text=postback_date+"は現在何時でも予約可能です"+chr(0x10008D))
             )
         else:
             line_bot_api.push_message(
                to=userId,
                messages=TextSendMessage(text=postback_date+"は"+str(reservedTime[0])+"以外なら空いています"+chr(0x10008D))
             )
         time_picker = TemplateSendMessage(
             alt_text='次に時間を設定してください',
             template=ButtonsTemplate(
                 text='以下から選択して下さい',
                 title='時間を選択して下さい。(予約可能なのは0,30分です)',
                 actions=[
                     DatetimePickerTemplateAction(
                         label='ここをタップ',
                         data='reservation2',
                         mode='time',
                         initial='10:00',
                         min='10:00',
                         max='21:00'
                     )
                 ]
             )
         )
         line_bot_api.push_message(
            to=userId,
            messages=time_picker
         )

    elif postback_msg=='reserveTime2':
         time_picker = TemplateSendMessage(
             alt_text='時間を変更してください',
             template=ButtonsTemplate(
                 text='以下から選択して下さい',
                 title='時間を選択して下さい。(予約可能なのは0,30分です)',
                 actions=[
                     DatetimePickerTemplateAction(
                         label='ここをタップ',
                         data='reservation2',
                         mode='time',
                         initial='10:00',
                         min='10:00',
                         max='21:00'
                     )
                 ]
             )
         )
         line_bot_api.push_message(
            to=userId,
            messages=time_picker
         )

#time confirmation
    elif postback_msg=='reservation2':
        #仮設定、DBから本当は取ってくる
        postback_date=wks.acell('D'+str(line_number[0])).value
        #postback_date='2020-06-10'
        postback_time=event.postback.params['time']
        wks.update_acell('E'+str(line_number[0]),postback_time)
        datetime_confirmation=ButtonsTemplate(
           title='予約日時確認',
           text=postback_date+" "+postback_time+'でお間違い無いでしょうか。',
           actions=[
               PostbackTemplateAction(
                   label='日にちを変更する',
                   data='changeDate'
               ),
               PostbackTemplateAction(
                   label='時間を変更する',
                   data='reserveTime2'
               ),
               PostbackTemplateAction(
                   label='予約最終ページへ',
                   data='LIFFpage'
               )

           ]
        )
        line_bot_api.push_message(
           to=userId,
           messages=TemplateSendMessage(
               alt_text='button template',
               template=datetime_confirmation
           )
        )

    elif postback_msg=='reservation3':
        postback_date=event.postback.params['date']
        wks.update_acell('D'+str(line_number[0]),postback_date)
        #postback_date=wks.acell('D'+str(line_number[0])).value
        #postback_date='2020-06-10'
        postback_time=wks.acell('E'+str(line_number[0])).value
        datetime_confirmation=ButtonsTemplate(
           title='予約日時確認',
           text=postback_date+" "+postback_time+'でお間違い無いでしょうか。',
           actions=[
               PostbackTemplateAction(
                   label='日にちを変更する',
                   data='changeDate'
               ),
               PostbackTemplateAction(
                   label='時間を変更する',
                   data='reserveTime'
               ),
               PostbackTemplateAction(
                   label='予約最終ページへ',
                   data='LIFFpage'
               )
           ]
        )
        line_bot_api.push_message(
           to=userId,
           messages=TemplateSendMessage(
               alt_text='button template',
               template=datetime_confirmation
           )
        )


    elif postback_msg=='LIFFpage':
        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="最後に名前、電話番号の入力をこちらから行ってください"+chr(0x10008D)
           +"\nhttps://liff.line.me/1654346616-2bKk0zXW"))


    elif postback_msg=='reserveComplete':
        line_bot_api.push_message(
           to=userId,
           messages=TextSendMessage(text="ご予約ありがとうございます"+chr(0x10008D)
           +"\n当日お待ちしております"+chr(0x10008F)))

    elif postback_msg=='changeReserve':
        postback_date=wks.acell('D'+str(line_number[0])).value
        postback_time=wks.acell('E'+str(line_number[0])).value
        postback_menu=wks.acell('B'+str(line_number[0])).value
        if postback_menu=="menu1":
            finalMenu="カット"
        elif postback_menu=="menu2":
            finalMenu="カラーリング"
        elif postback_menu=="menu3":
            finalMenu="有機デジタルパーマ"
        elif postback_menu=="menu4":
            finalMenu="ヘッドスパ"
        changeReservation=ButtonsTemplate(
           #title='予約最終確認',
           text='現在の予約内容は\n・お名前: '+finalName+"・日時: "+postback_date+" "+postback_time+"\n・メニュー: "
                    +finalMenu +"です。",
           actions=[
               PostbackTemplateAction(
                   label='変更しない',
                   data='reserveComplete'
               ),
               PostbackTemplateAction(
                   label='日にちを変更する',
                   data='changeDate'
               ),
               PostbackTemplateAction(
                   label='時間を変更する',
                   data='reserveTime2'
               ),
               PostbackTemplateAction(
                   label='メニューを変更する',
                   data='reserveComplete'
               ),
           ]
        )
        line_bot_api.push_message(
           to=userId,
           messages=TemplateSendMessage(
               alt_text='button template',
               template=reserve_confirmation
           )
        )

    elif postback_msg=='reserveComplete':
        response_json_list=[]
        result_dict={
            "thumbnail_image_url": "https://imgbp.hotp.jp/CSP/IMG_SRC/55/12/B060655512/B060655512_271-361.jpg",
            "title": "カット",
            "text": "カットにシャンプー・ブロー込のコースです！",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu1_"
                }}
        response_json_list.append(result_dict)
        result_dict2={
            "thumbnail_image_url": "https://cloudinary-a.akamaihd.net/vivivi/image/upload/t_beauty,f_auto,dpr_2.0,q_auto:good/c_pad,w_500,h_500/v1574838084/stl_43760_5dde1f433079f_01.jpg",
            "title": "カラーリング",
            "text": "トリートメントつきカラーリングコースです！長さ一律料金です",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu2_"
                }}
        response_json_list.append(result_dict2)
        result_dict3={
            "thumbnail_image_url": "https://d3kszy5ca3yqvh.cloudfront.net/wp-content/uploads/2019/3/29/16/083d03f7740809c632d20f5e9d6afc86.jpg",
            "title": "有機デジタルパーマ",
            "text": "持ちが良く髪に負担をかけにくい有機デジタルパーマのコースです！",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu3_"
                }}
        response_json_list.append(result_dict3)
        result_dict4={
            "thumbnail_image_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcRtkxvQ8nQ9-DgHCmcgPw23g_4mJzjxOwvH49c9IIbrqLP-YV0J&usqp=CAU",
            "title": "ヘッドスパ",
            "text": "自分へのご褒美に！ヘッドスパ、有機炭酸シャンプー、トリートメントのコース",
            "actions": {
                "label": "このメニューで決定する",
                "data":"menu4_"
                }}
        response_json_list.append(result_dict4)

        columns = [
            CarouselColumn(
                thumbnail_image_url=column["thumbnail_image_url"],
                title=column["title"],
                text=column["text"],
                actions=[
                PostbackTemplateAction(
                    label=column["actions"]["label"],
                    data=column["actions"]["data"]
                ),
            ])
            for column in response_json_list
        ]
        messages = TemplateSendMessage(
            alt_text="メニュー選択",
            template=CarouselTemplate(columns=columns),
        )
        print("messages is: {}".format(messages))

        line_bot_api.reply_message(
            event.reply_token,
            messages=messages
        )










@handler.add(FollowEvent)
def handle_follow(event):
    line_bot_api.reply_message(
        event.reply_token, TextSendMessage(text=FOLLOWED_MESSAGE)
    )


@handler.add(UnfollowEvent)
def handle_unfollow():
    app.logger.info("Got Unfollow event")

#-------休みをカレンダーから取得--------
SCOPES = ['https://www.googleapis.com/auth/calendar']
#SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def getHoliday(num):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            #secret = json.loads(os.environ[{"GOOGLE_CALENDAR_API_KEY"}])
            # credential = {
            #     "type": "service_account",
            #     "project_id": os.environ['PROJECT_ID'],
            #     "private_key_id": os.environ['PRIVATE_KEY_ID'],
            #     "private_key": os.environ['PRIVATE_KEY'],
            #     "client_email": os.environ['CLIENT_SECRET'],
            #     "client_id": os.environ['CLIENT_ID'],
            #     "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            #     "token_uri": "https://oauth2.googleapis.com/token",
            #     "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            #     "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/samplepra%40quickstart-1591975098729.iam.gserviceaccount.com"
            #  }
            #flow = ServiceAccountCredentials.from_json_keyfile_dict(credential,SCOPES)
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    today=datetime.date.today()
    laterday=today + datetime.timedelta(days=14)
    #print(today,laterday)
    today=today.strftime('%Y/%m/%d')
    laterday=laterday.strftime('%Y/%m/%d')
    today = datetime.datetime.strptime(today, '%Y/%m/%d').isoformat()+'Z'
    laterday = datetime.datetime.strptime(laterday, '%Y/%m/%d').isoformat()+'Z'
    _carenderID={"1":'amreo6lnsdt9jai39pepj3kb18@group.calendar.google.com',"2":"k12ajhtfaemki7s7tuc8setfdk@group.calendar.google.com"}
    carenderID=_carenderID[str(num)]

    #events_result = service.events().list(calendarId='amreo6lnsdt9jai39pepj3kb18@group.calendar.google.com',
    events_result = service.events().list(calendarId=carenderID,
                                        timeMin=today,
                                        timeMax=laterday,
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    holiday=[]
    if not events:
        print('No upcoming events found.')
    for event in events:
        if event['summary']=="休み":
            day = event['start'].get('dateTime', event['start'].get('date'))
            day=day[5:10]
            holiday.append(day)


    return  holiday

def getTime(stylistNum,date):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    _date=date.replace("-","/")
    timeto=datetime.datetime.strptime(_date, '%Y/%m/%d')
    timeto= str(timeto + datetime.timedelta(days=1))
    timeto=timeto.replace(" 00:00:00","").replace("-","/")

    timefrom = datetime.datetime.strptime(_date, '%Y/%m/%d').isoformat()+'Z'
    timeto= datetime.datetime.strptime(timeto, '%Y/%m/%d').isoformat()+'Z'
    _carenderID={"stylist1":'amreo6lnsdt9jai39pepj3kb18@group.calendar.google.com',"stylist2":"k12ajhtfaemki7s7tuc8setfdk@group.calendar.google.com"}
    carenderID=_carenderID[stylistNum]
    events_result = service.events().list(calendarId=carenderID,
                                        timeMin=timefrom,
                                        timeMax=timeto,
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        times="None"
    else:
        times=[]
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            end = event['end'].get('dateTime', event['end'].get('date'))
            start = datetime.datetime.strptime(start[:-6], '%Y-%m-%dT%H:%M:%S').strftime('%H:%M')
            end = datetime.datetime.strptime(end[:-6], '%Y-%m-%dT%H:%M:%S').strftime('%H:%M')
            #print(start,end)
            time=start+"〜"+end
            times.append(time)


    return times





#--------googleカレンダーに予約書き込み----
def printCalendar(userId,g_name,g_date,g_time,g_menu):
    scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('sheets.json', scope)
    gc = gspread.authorize(credentials)
    wks = gc.open('美容院予約 bot sample').sheet1
    line_number=[]
    for i in range(2,len(wks.col_values(1))+2):
        if userId==wks.acell('A'+str(i)).value:
            wks.update_acell('I'+str(i),"selected")
            line_number.append(i)
            break
        elif wks.acell('A'+str(i)).value=="":
            wks.update_acell('A'+str(i),userId)
            wks.update_acell('I'+str(i),"selected")
            line_number.append(i)
            break
    stylistNum=wks.acell('C'+str(line_number[0])).value
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token1.pickle'):
        with open('token1.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token1.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    dateTime=g_date+"T"+g_time+":00"
    _lastTime=g_date+" "+g_time
    lastTime = datetime.datetime.strptime(_lastTime, '%Y-%m-%d %H:%M')
    lt = lastTime + datetime.timedelta(hours=1, minutes=30)
    ltime= lt.strftime("%Y-%m-%d %H:%M")
    lasttim=ltime.replace(" ","T")
    lasttime=lasttim+":00"


    # Call the Calendar API
    event = {
      'summary': g_name,
      'description': g_menu,
      'start': {
        'dateTime': dateTime,
        'timeZone': 'Japan',
      },
      'end': {
        'dateTime': lasttime,
        'timeZone': 'Japan',
      },
    }
    _carenderID={"stylist1":'amreo6lnsdt9jai39pepj3kb18@group.calendar.google.com',"stylist2":"k12ajhtfaemki7s7tuc8setfdk@group.calendar.google.com"}
    #carenderID=_carenderID[str(stylistNum)]
    carenderID=_carenderID[stylistNum]
    event = service.events().insert(calendarId=carenderID,
                                    body=event).execute()




if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
