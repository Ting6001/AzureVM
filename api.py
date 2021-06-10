from flask import Flask, request
import random

app = Flask(__name__)
app.config['DEBUG'] = True

@app.route('/', methods=['GET'])
def Home():
    if 'name' in request.args:
        name = request.args['name']
        return "Hello {}!".format(name)
    return "<h1> Hello Flask!</h1>"

@app.route('/api', methods=['GET'])
def Getworkrate():
    dic = {'230000':[], '23R000':[]}
        
    dic_DivInfo = { '230000':{  'DEP':['230220', '230R30'],
                                'SUB':['ME']
                            }, 
                    '23R000':{  'DEP':['23R000', '23R200', '23R300'],
                                'SUB':['ME', 'HW']
                        }}
    for div, data in dic_DivInfo.items():
        for key, lst in data.items():
            for item in lst:
                print(div, key, item)
                dic_tmp = {}
                dic_tmp['division'] = div
                dic_tmp['type'] = key
                dic_tmp['title'] = item
                dic_tmp['HC'] = 0
                for i in range(1,4):
                    dic_tmp['b_'+str(i)] = round(random.uniform(0.8, 1.3),1)
                    dic_tmp['a_'+str(i)] = round(random.uniform(0.8, 1.3),1)  
                dic[div].append(dic_tmp)

    if 'name' in request.args:
        name = request.args['name']
        return "<h1> Hello {}!</h1>".format(name)
    elif 'div' in request.args:
        div_pa = request.args['div']
        return {'div':div_pa,
                'data':dic.get(div_pa, {})}
    else:
        return "Error: No storeid provided. Please specify a storeid."


import os
if __name__ == "__main__":
    port = os.environ.get('PORT',5000)
    app.run(host='0.0.0.0',port=port)