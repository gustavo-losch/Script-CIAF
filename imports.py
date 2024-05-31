import tabula
import pyautogui
import time
from datetime import *
import tkinter
import customtkinter
import re
from dbfread import DBF
import pandas as pd
import openpyxl
import sys
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, BaseDocTemplate, PageTemplate, Frame, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing, Line
from PIL import Image, ImageTk
import shutil
import csv
import tkPDFViewer
import pymupdf