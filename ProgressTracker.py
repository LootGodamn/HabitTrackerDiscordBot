import asyncio
from asyncio.windows_events import NULL
import random
from typing import List
import discord
from discord.ext import commands
import openpyxl

bot = commands.Bot(command_prefix='#', help_command=None)

@bot.command()
@commands.adduser(1.0, 10.0, commands.BucketType.user)
async def showlist(ctx):
    await typing(ctx)
    #Take username -> search in table 
    # Y -> return alr exist error
    # N -> add to table + add new sheet for user w template

@bot.command()
@commands.addtask(1.0, 10.0, commands.BucketType.user)
async def showlist(ctx):
    await typing(ctx)
    #Take username + task name + Duration + Difficulty + desc 
    #Validate -> add to corresponding sheet

@bot.command()
@commands.removetask(1.0, 10.0, commands.BucketType.user)
async def showlist(ctx):
    await typing(ctx)
    #Take username + Task name
    #search in corresponding sheet 
    #Y -> remove
    #N -> error

@bot.command()
@commands.setdailygoal(1.0, 10.0, commands.BucketType.user)
async def showlist(ctx):
    await typing(ctx)
    #Take username + amnt
    #change value in corresponding sheet

@bot.command()
@commands.finishtask(1.0, 10.0, commands.BucketType.user)
async def showlist(ctx):
    await typing(ctx)
    #Take username + Task name
    #search in corresponding sheet 
    #Y -> remove row -> increment values
    #N -> error, use addtask

@bot.command()
@commands.showtasks(1.0, 10.0, commands.BucketType.user)
async def showlist(ctx):
    await typing(ctx)
    #loop thru each row 
        #loop thru each column
            #concat values
        #join 
    #display

bot.run('')
# # Don't reveal your bot token, regenerate it asap if you do