import asyncio
import logging
from bot import start_bot

logging.basicConfig(filename='info.log',
                    filemode='w',
                    level=logging.INFO,
                    format='%(name)s - %(levelname)s - %(message)s')


async def main():
    await start_bot()


if __name__ == '__main__':
    asyncio.run(main())
