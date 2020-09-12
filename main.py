import telebot
from telebot import types
import openpyxl as opx
from openpyxl import cell
# from config import userProfile

workbook = opx.load_workbook(filename = r'D:\User Files\Рабочий стол\timetable2.xlsx', data_only=True)
sheet = workbook.get_sheet_by_name('Лист1')

bot = telebot.TeleBot('')

userProfile = ''
toSend = ''

n = 1
for cellObj in sheet['B'+str(n):'B61']:
		for cell in cellObj:
			if cell.value == 'Понедельник':
				print(cell.value, str(n))
				monday = str(n)

			elif cell.value == 'Вторник':
				print(cell.value, str(n))
				tuesday = str(n)

			elif cell.value == 'Среда':
				print(cell.value, str(n))
				wednesday = str(n)

			elif cell.value == 'Четверг':
				print(cell.value, str(n))
				thursday = str(n)

			elif cell.value == 'Пятница':
				print(cell.value, str(n))
				friday = str(n)

			elif cell.value == 'Суббота':
				print(cell.value, str(n))
				saturday = str(n)
			n += 1


##############################################################################################################


@bot.message_handler(commands=['start'])
def show_profiles(message):
	bot.send_message(message.chat.id, 'С помощью меня ты сможешь посмотреть расписание любого профиля на любой интересующий тебя день.\nПожелания, жалобы - @fransson.')

	markup_profile = types.InlineKeyboardMarkup()
	item_Gum = types.InlineKeyboardButton(text = 'Гуманитарный', callback_data = 'gum')
	item_Himbio = types.InlineKeyboardButton(text = 'Химико-Биологический', callback_data = 'himbio')
	item_Fizmat = types.InlineKeyboardButton(text = 'Физико-Математический', callback_data = 'fizmat')
	
	markup_profile.add(item_Gum, item_Himbio, item_Fizmat)
	bot.send_message(message.chat.id, 'Выбери профиль', reply_markup = markup_profile)


@bot.callback_query_handler(func = lambda call: True)
def getting_call(call):
	# ВЫБОР ПРОФИЛЯ
	global userProfile
	global toSend

	if call.data == 'gum':
		userProfile = 'Гуманитарный'
		bot.answer_callback_query(call.id, text=f'Вы выбрали {userProfile} профиль')

		show_keys(call)


	elif call.data == 'himbio':
		userProfile = 'Химико-Биологический'
		bot.answer_callback_query(call.id, text=f'Вы выбрали {userProfile} профиль')

		show_keys(call)


	elif call.data == 'fizmat':
		userProfile = 'Физико-Математический'
		bot.answer_callback_query(call.id, text=f'Вы выбрали {userProfile} профиль')

		show_keys(call)
	

	elif call.data == 'monday':
		toSend = ''
		try:
			if userProfile == 'Гуманитарный':
				for cellObj in sheet['AA'+monday:'AA'+str(int(monday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Химико-Биологический':
				for cellObj in sheet['AC'+monday:'AC'+str(int(monday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Физико-Математический':
				for cellObj in sheet['AE'+monday:'AE'+str(int(monday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			bot.send_message(call.message.chat.id, f'ПОНЕДЕЛЬНИК\n\n{toSend}')

		except NameError:
			bot.answer_callback_query(call.id, text='Нет данных. Возможно, расписание на данный день еще не составлено.')


	elif call.data == 'tuesday':
		toSend = ''
		try:
			if userProfile == 'Гуманитарный':
				for cellObj in sheet['AA'+tuesday:'AA'+str(int(tuesday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Химико-Биологический':
				for cellObj in sheet['AC'+tuesday:'AC'+str(int(tuesday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Физико-Математический':
				for cellObj in sheet['AE'+tuesday:'AE'+str(int(tuesday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			bot.send_message(call.message.chat.id, f'ВТОРНИК\n\n{toSend}')

		except NameError:
			bot.answer_callback_query(call.id, text='Нет данных. Возможно, расписание на данный день еще не составлено.')
			

	elif call.data == 'wednesday':
		toSend = ''
		try:
			if userProfile == 'Гуманитарный':
				for cellObj in sheet['AA'+wednesday:'AA'+str(int(wednesday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Химико-Биологический':
				for cellObj in sheet['AC'+wednesday:'AC'+str(int(wednesday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Физико-Математический':
				for cellObj in sheet['AE'+wednesday:'AE'+str(int(wednesday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			bot.send_message(call.message.chat.id, f'СРЕДА\n\n{toSend}')

		except NameError:
			bot.answer_callback_query(call.id, text='Нет данных. Возможно, расписание на данный день еще не составлено.')


	elif call.data == 'thursday':
		toSend = ''
		try:
			if userProfile == 'Гуманитарный':
				for cellObj in sheet['AA'+thursday:'AA'+str(int(thursday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Химико-Биологический':
				for cellObj in sheet['AC'+thursday:'AC'+str(int(thursday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Физико-Математический':
				for cellObj in sheet['AE'+thursday:'AE'+str(int(thursday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			bot.send_message(call.message.chat.id, f'ЧЕТВЕРГ\n\n{toSend}')

		except NameError:
			bot.answer_callback_query(call.id, text='Нет данных. Возможно, расписание на данный день еще не составлено.')


	elif call.data == 'friday':
		toSend = ''
		try:
			if userProfile == 'Гуманитарный':
				for cellObj in sheet['AA'+friday:'AA'+str(int(friday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Химико-Биологический':
				for cellObj in sheet['AC'+friday:'AC'+str(int(friday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Физико-Математический':
				for cellObj in sheet['AE'+friday:'AE'+str(int(friday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			bot.send_message(call.message.chat.id, f'ПЯТНИЦА\n\n{toSend}')

		except NameError:
			bot.answer_callback_query(call.id, text='Нет данных. Возможно, расписание на данный день еще не составлено.')


	elif call.data == 'saturday':
		toSend = ''
		try:
			if userProfile == 'Гуманитарный':
				for cellObj in sheet['AA'+saturday:'AA'+str(int(saturday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			elif userProfile == 'Химико-Биологический':
				for cellObj in sheet['AC'+saturday:'AC'+str(int(saturday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'	

			elif userProfile == 'Физико-Математический':
				for cellObj in sheet['AE'+saturday:'AE'+str(int(saturday)+8)]:
					for cell in cellObj:
						if cell.value != None:
							toSend += cell.value + '\n'

			bot.send_message(call.message.chat.id, f'СУББОТА\n\n{toSend}')

		except NameError:
			bot.answer_callback_query(call.id, text='Нет данных. Возможно, расписание на данный день еще не составлено.')

def show_keys(call):
	markup_day = types.InlineKeyboardMarkup()
	item_monday = types.InlineKeyboardButton(text = 'ПН', callback_data = 'monday')
	item_tuesday = types.InlineKeyboardButton(text = 'ВТ', callback_data = 'tuesday')
	item_wednesday = types.InlineKeyboardButton(text = 'СР', callback_data = 'wednesday')
	item_thursday = types.InlineKeyboardButton(text = 'ЧТ', callback_data = 'thursday')
	item_friday = types.InlineKeyboardButton(text = 'ПТ', callback_data = 'friday')
	item_saturday = types.InlineKeyboardButton(text = 'СБ', callback_data = 'saturday')
	
	markup_day.add(item_monday, item_tuesday, item_wednesday, item_thursday, item_friday, item_saturday)
	bot.send_message(call.message.chat.id, 'Выбери день', reply_markup = markup_day)

bot.polling(none_stop=True)
