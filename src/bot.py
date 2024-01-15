from aiogram import Bot, Dispatcher, Router
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext

from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, FSInputFile
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import StateFilter

from config import API_TOKEN_BOT
from excel_handler import clear_directory, process_excels

router = Router()


class FileReports(StatesGroup):
    files = State()


class FileReport(StatesGroup):
    file = State()


@router.message(Command('load'))
async def load(message: Message, state: FSMContext):
    await message.answer('Пришлите пожалуйста файлы отчетов в формате .xls или .xlxs\n'
                         'По завершению напишите мне слово <code>Стоп</code>', parse_mode='html')
    clear_directory('reports')
    clear_directory('ready_report')
    await state.set_state(FileReports.files)


@router.message(Command('load1'))
async def load(message: Message, state: FSMContext):
    await message.answer('Пришлите пожалуйста один фай отчета в формате .xls или .xlxs', parse_mode='html')
    clear_directory('reports_all')
    clear_directory('ready_report')
    await state.set_state(FileReport.file)


@router.message(StateFilter(FileReports.files))
async def accepting_files(message: Message, state: FSMContext):
    if message.text and message.text.lower() == 'стоп':
        await state.clear()
        status = process_excels('reports')
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
            file = await message.bot.get_file(file_id)
            file_path = file.file_path

            await message.bot.download_file(file_path, f"reports/{file_path[10:]}")
            await message.answer("Файл Excel успешно получен и сохранен.")
        else:
            await message.answer('Вы прислали файл не того формата\n'
                                 'Пожалуйста, отправьте файл формата .xlsx или .xls.')


@router.message(StateFilter(FileReport.file))
async def accepting_file(message: Message, state: FSMContext):
    await state.clear()

    if message.document:
        mime_type = message.document.mime_type
        file_id = message.document.file_id

        if mime_type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'application/vnd.ms-excel']:
            file = await message.bot.get_file(file_id)
            file_path = file.file_path

            await message.bot.download_file(file_path, f"reports_all/{file_path[10:]}")
        else:
            await message.answer('Вы прислали файл не того формата\n'
                                 'Пожалуйста, отправьте файл формата .xlsx или .xls.')
    else:
        await message.answer("Вы прислали не файл excel")
        return

    status = process_excels('reports_all')
    if status == 'файл создан':
        file = FSInputFile('ready_report/report.xlsx', filename='report.xlsx')
        await message.bot.send_document(chat_id=message.from_user.id, document=file,
                                        caption='Сгенерированный отчет по файлам')
    else:
        await message.answer(status)
    return


async def start_bot():
    bot = Bot(token=API_TOKEN_BOT)
    storage = MemoryStorage()
    dp = Dispatcher(bot=bot, storage=storage)
    dp.include_router(router)
    await dp.start_polling(bot)
