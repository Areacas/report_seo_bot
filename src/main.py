import asyncio
import logging
from bot import start_bot
import os

logging.basicConfig(filename='info.log',
                    filemode='w',
                    level=logging.INFO,
                    format='%(name)s - %(levelname)s - %(message)s')


directories = ["reports", "ready_report"]
for directory in directories:
    if not os.path.exists(directory):
        os.makedirs(directory)


async def main():
    await start_bot()


if __name__ == '__main__':
    asyncio.run(main())
