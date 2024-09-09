# Initial imports
from datetime import datetime, timezone
import time
from telethon.sync import TelegramClient
from openpyxl import Workbook
import asyncio

# Setup / change only the first time you use it
username = '+...'  # Telegram account name
phone = '+...'  # Telegram phone number
api_id = '...'  # API ID from https://my.telegram.org/apps
api_hash = '...'  # API hash from https://my.telegram.org/apps
url = ''
index = 1

# Setup / change every time to use to define scraping
channel = '@xxxx'  # The name of the channel or group to scrape
file_path = 'xxxx.xlsx'  # File path to save the XLSX file
d_min = xx  # Start day (inclusive)
m_min = x  # Start month
y_min = xxxx  # Start year
d_max = xx  # End day (exclusive)
m_max = x  # End month
y_max = xxxx  # End year
key_search = ''  # Keyword for search (if any)

# Create a new workbook and select the active worksheet
wb = Workbook()
ws = wb.active
ws.title = "Telegram Data"

# titles and workspace
titles = ['Scraping ID', 'Group', 'Author ID', 'Content', 'Date', 'Message ID', 'Author', 'Views', 'Reactions', 'Shares', 'Media', 'Comments']
ws.append(titles)

async def scrape_telegram():
    global index
    scraped_data = []
    start_date = datetime(y_min, m_min, d_min, tzinfo=timezone.utc)
    end_date = datetime(y_max, m_max, d_max, tzinfo=timezone.utc)
    
    async with TelegramClient(username, api_id, api_hash) as client:
        async for message in client.iter_messages(channel, search=key_search):
            print(message.date)
            if message.date < start_date:
                break  # Exit the loop if the message date is earlier than the start date
            if start_date <= message.date < end_date:
                # Check if there is media
                url = f'https://t.me/{channel}/{message.id}'.replace('@', '') if message.media else 'no media'

                # Check if there are reactions
                emoji_string = ""
                if message.reactions:
                    for reaction_count in message.reactions.results:
                        emoji = reaction_count.reaction.emoticon
                        count = str(reaction_count.count)
                        emoji_string += f"{emoji} {count} "

                # Content condensation
                content = [f'#ID{index:05}', channel, message.sender_id, message.text, message.date.strftime('%Y-%m-%d %H:%M:%S'), message.id, message.post_author, message.views, emoji_string, message.forwards, url]

                # Collect comments
                comments = []
                try:
                    async for reply in client.iter_messages(channel, reply_to=message.id):
                        comments.append(reply.text)
                except:
                    comments = ['possible adjustment']
                comments = ';\n'.join(comments)

                # Append comments to content
                content.append(comments)

                # Append content to scraped_data
                scraped_data.append(content)

                # Update progress counter
                print(f'Item {index:05} completed!')
                print(f'Id: {message.id:05}.\n')

                # Update loop parameters
                index += 1

    # Write all scraped data to the worksheet in one go
    for row in scraped_data:
        ws.append(row)

# Run the async function
asyncio.run(scrape_telegram())

# Save the workbook
wb.save(file_path)

# End
print(f'----------------------------------------\n#Concluded! #{index-1:05} posts were scraped!\n----------------------------------------\n\n\n\n')
