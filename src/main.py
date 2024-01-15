import asyncio
import logging
from bot import start_bot
import os


async def main():
    logging.basicConfig(filename='info.log',
                        filemode='w',
                        level=logging.INFO,
                        format='%(name)s - %(levelname)s - %(message)s')

    if not os.path.exists("reports"):
        os.makedirs("reports")

    await start_bot()


if __name__ == '__main__':
    asyncio.run(main())
