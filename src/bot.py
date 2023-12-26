from aiogram import Bot, Dispatcher
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext

from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, FSInputFile
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import StateFilter

from config import API_TOKEN_BOT
from excel_handler import clear_directory, process_excels

bot = Bot(token=API_TOKEN_BOT)
storage = MemoryStorage()
dp = Dispatcher(bot=bot, storage=storage)


class FileReports(StatesGroup):
    files = State()


@dp.message(Command('load'))
async def load(message: Message, state: FSMContext):
    await message.answer('Пришлите пожалуйста файлы отчетов в формате .xls или .xlxs\n'
                         'По завершению напишите мне слово <code>Стоп</code>', parse_mode='html')
    clear_directory('reports')
    clear_directory('ready_report')
    await state.set_state(FileReports.files)


@dp.message(StateFilter(FileReports.files))
async def accepting_files(message: Message, state: FSMContext):
    if message.text and message.text.lower() == 'стоп':
        await state.clear()
        status = process_excels()
        if status == 'файл создан':
            file = FSInputFile('ready_report/report.xlsx', filename='report.xlsx')
            await message.bot.send_document(chat_id=message.from_user.id, document=file,
                                            caption='Сгенерированный отчет по файлам')
        else:
            await message.answer(status)
        return

    if message.document:
        mime_type = message.document.mime_type
        file_id = message.document.file_id

        if mime_type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'application/vnd.ms-excel']:
            file = await bot.get_file(file_id)
            file_path = file.file_path

            await bot.download_file(file_path, f"reports/{file_path[10:]}")
            await message.answer("Файл Excel успешно получен и сохранен.")
        else:
            await message.answer('Вы прислали файл не того формата\n'
                                 'Пожалуйста, отправьте файл формата .xlsx или .xls.')


async def start_bot():
    await dp.start_polling(bot)
