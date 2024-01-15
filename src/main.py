import asyncio
import logging
from bot import start_bot
import os


async def main():
    logging.basicConfig(filename='info.log',
                        filemode='w',
                        level=logging.INFO,
                        format='%(name)s - %(levelname)s - %(message)s')

    directories = ["reports", "reports_all", "ready_report"]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

    await start_bot()


if __name__ == '__main__':
    asyncio.run(main())
