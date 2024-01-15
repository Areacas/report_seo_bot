from aiogram import Bot, Dispatcher, Router
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext

from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, FSInputFile
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import StateFilter

from config import API_TOKEN_BOT
from excel_handler import clear_directory, process_excels
from str import creat_directories

router = Router()


class FileReportsSku(StatesGroup):
    files = State()


class FileReportGroup(StatesGroup):
    files = State()


@router.message(Command('load_sku'))
async def load(message: Message, state: FSMContext):
    await message.answer('Пришлите файлы отчетов по <b>sku</b> в формате .xls или .xlxs\n'
                         'По завершению напишите мне слово:\n<code>Стоп</code>', parse_mode='html')
    user_id = message.from_user.id
    creat_directories(user_id)
    clear_directory(f'reports/{user_id}/reports_sku')
    clear_directory(f'reports/{user_id}/ready_reports_sku')
    await state.set_state(FileReportsSku.files)


@router.message(Command('load_group'))
async def load(message: Message, state: FSMContext):
    await message.answer('Пришлите файлы отчетов по <b>group</b> в формате .xls или .xlxs\n'
                         'По завершению напишите мне слово:\n<code>Стоп</code>', parse_mode='html')
    user_id = message.from_user.id
    creat_directories(user_id)
    clear_directory(f'reports/{user_id}/reports_group')
    clear_directory(f'reports/{user_id}/ready_reports_group')
    await state.set_state(FileReportGroup.files)


@router.message(StateFilter(FileReportsSku.files))
async def accepting_files(message: Message, state: FSMContext):
    user_id = message.from_user.id
    if message.text and message.text.lower() == 'стоп':
        await state.clear()
        status = process_excels('reports_sku', user_id)
        if status == 'файл создан':
            file = FSInputFile(f'reports/{user_id}/ready_reports_sku/report.xlsx', filename='report.xlsx')
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

            await message.bot.download_file(file_path, f"reports/{user_id}/reports_sku/{file_path[10:]}")
            await message.answer("Файл Excel успешно получен и сохранен.")
        else:
            await message.answer('Вы прислали файл не того формата\n'
                                 'Пожалуйста, отправьте файл формата .xlsx или .xls.')


@router.message(StateFilter(FileReportGroup.files))
async def accepting_file(message: Message, state: FSMContext):
    user_id = message.from_user.id
    if message.text and message.text.lower() == 'стоп':
        await state.clear()
        status = process_excels('reports_group', user_id)
        if status == 'файл создан':
            file = FSInputFile(f'reports/{user_id}/ready_reports_group/report.xlsx', filename='report.xlsx')
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

            await message.bot.download_file(file_path, f"reports/{user_id}/reports_group/{file_path[10:]}")
            await message.answer("Файл Excel успешно получен и сохранен.")
        else:
            await message.answer('Вы прислали файл не того формата\n'
                                 'Пожалуйста, отправьте файл формата .xlsx или .xls.')
    else:
        await message.answer("Вы прислали не файл excel")
        return


async def start_bot():
    bot = Bot(token=API_TOKEN_BOT)
    storage = MemoryStorage()
    dp = Dispatcher(bot=bot, storage=storage)
    dp.include_router(router)
    await dp.start_polling(bot)
