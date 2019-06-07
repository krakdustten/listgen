from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import workbookWriter as wrt
from flask_apscheduler import APScheduler
from os import walk, path, remove
import datetime
from enum import Enum

app = Flask(__name__)
CORS(app)
scheduler = APScheduler()
scheduler.init_app(app)
scheduler.start()


class Config(object):
    SCHEDULER_API_ENABLED = True


class colType(Enum):
    AMOUNT_NEEDED = (1, "Amount needed")
    PRICE_PER_PIECE = (2, "price/p")
    MIN_AMOUNT = (3, "Min amount")
    AMOUNT = (4, "Amount")
    PRICE = (5, "price")
    NONE = (6, "ERROR")

    def __init__(self, number, colname):
        self.number = number
        self.colname = colname

    @staticmethod
    def from_str(label):
        if label in ('amount needed', 'an'):
            return colType.AMOUNT_NEEDED
        if label in ('price per piece', 'ppp'):
            return colType.PRICE_PER_PIECE
        if label in ('min amount', 'mamt'):
            return colType.MIN_AMOUNT
        if label in ('amount', 'amt'):
            return colType.AMOUNT
        if label in ('price', 'p'):
            return colType.PRICE
        return colType.NONE


@app.route('/', methods=['POST'])
def hello_world():
    data = request.get_json(silent=False)
    config = {
        'fileType': wrt.file.XLSX,
        'fileName': "file",
        'numFormat': '€#,##'
    }
    header = {
        'titles': [],
        'type': []
    }
    cells = []
    #Get configuration from json
    if 'config' in data:
        data_config = data['config']
        if 'fileType' in data_config:
            fileType = wrt.file.from_str(data_config['fileType'])
            if fileType is None:
                return jsonify(error="FileType not found.")
            config['fileType'] = fileType
        if 'fileName' in data_config:
            config['fileName'] = data_config['fileName']
        if 'numFormat' in data_config:
            config['numFormat'] = data_config['numFormat']
    #Get header from json
    if 'header' not in data:
        return jsonify(error="No header.")
    data_header = data['header']
    if 'titles' in data_header:
        data_titles = data_header['titles']
        if isinstance(data_titles, list):
            for dt in data_titles:
                header['titles'].append(dt)
        else:
            header['titles'].append(data_titles)
    if 'columnType' in data_header:
        data_columnTypes = data_header['columnType']
        if isinstance(data_columnTypes, list):
            for dt in data_columnTypes:
                header['type'].append(colType.from_str(dt))
        else:
            header['type'].append(colType.from_str(data_columnTypes))
    new_titles = []
    new_type = []
    ColNames = {}
    for i in range(max(len(header['titles']), len(header['type']))):
        if i < len(header['titles']):
            title = header['titles'][i]
        else:
            title = ""
        if i < len(header['type']):
            type = header['type'][i]
            if type == colType.AMOUNT_NEEDED:
                ColNames['amountNeeded'] = wrt.intToCol(i)
            if type == colType.PRICE_PER_PIECE:
                ColNames['costPP'] = wrt.intToCol(i)
            if type == colType.MIN_AMOUNT:
                ColNames['minAmount'] = wrt.intToCol(i)
            if type == colType.AMOUNT:
                ColNames['amount'] = wrt.intToCol(i)
            if type == colType.PRICE:
                ColNames['price'] = wrt.intToCol(i)
        else:
            type = colType.NONE
        if title == "":
            title = type.colname
        new_titles.append(title)
        new_type.append(type)
    header['titles'] = new_titles
    header['type'] = new_type
    #Get data from json
    if 'data' not in data:
        return jsonify(error="No data.")
    data_data = data['data']
    collength = len(header['titles'])
    if isinstance(data_data, list):
        for row in data_data:
            if isinstance(row, list):
                crow = []
                for cell in row:
                    crow.append(cell)
                collength = max(collength, len(crow))
            else:
                return jsonify(error="Error in data.")
            cells.append(crow)
    else:
        return jsonify(error="Error in data.")
    rowlength = len(cells)
    for row in cells:
        while len(row) < collength:
            row.append("")
    #Write to the file
    t = wrt.workbookWriter(config['fileType'], config['fileName'], config['numFormat'])
    for i in range(rowlength):
        row = cells[i]
        for j in range(collength):
            cell = row[j]
            populateCell(t, cell, header['type'][j], i + 1, j, ColNames)
    for i in range(collength):
        t.writeCell(i, 0, header['titles'][i], t.format_bold)
    t.close()

    return send_file(t.fileName, as_attachment=True, attachment_filename=(t.name + t.fileType.extension))


def populateCell(t, cell, type, y, x, ColNames):
    if cell != "":
        t.writeCell(x, y, cell)
        return
    if type == colType.AMOUNT:
        if ('amountNeeded' in ColNames) and ('minAmount' in ColNames):
            t.writeCellFormula(x, y, "=CEILING(" + ColNames['amountNeeded'] + str(y + 1) + "," + ColNames['minAmount'] + str(y + 1) + ")")
    if type == colType.PRICE:
        if ('costPP' in ColNames) and ('amount' in ColNames):
            t.writeCellFormula(x, y, "=" + ColNames['amount'] + str(y + 1) + "*" + ColNames['costPP'] + str(y + 1), t.format_money)
    return


@scheduler.task('interval', id='do_job_1', minutes=1, misfire_grace_time=900)
def backgroundSchedual():
    now = datetime.datetime.now()
    files = []
    for (dirpath, dirnames, filenames) in walk("files/"):
        files.extend(filenames)
        break
    for f in files:
        time = path.getmtime("files/" + f)
        if datetime.datetime.fromtimestamp(time) < now - datetime.timedelta(hours=1):
            remove("files/" + f)


backgroundSchedual()
if __name__ == '__main__':
    app.run()




