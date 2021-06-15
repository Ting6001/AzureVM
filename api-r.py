from flask import Flask, request 
from flask_restful import Resource, Api,reqparse
# from werkzeug.contrib.fixers import ProxyFix # new
import random

app = Flask(__name__)
app.config['DEBUG'] = True
api = Api(app)

class WorkRate (Resource):
    parser = reqparse.RequestParser()
    # parser.add_argument('email', required=True, help='Email is required')
    # parser = parser.add_argument('lst', type=str, location='json', action="append")
    parser = parser.add_argument('data', type=dict, location='json', action="append", help='data dict is required')

    def post(self): # create
        arg = self.parser.parse_args()
        # print(arg)
        set_dep = set()
        set_sub = set()

        for item in arg['data']:
            if item['HC'] > 0:
                set_dep.add(item['deptid'])
                set_sub.add(item['sub_job_family'])
        div = arg['data'][0]['div']
        print('div:', div)
        print('dep:', set_dep)
        print('sub:', set_sub)

        lst_return = []    
        dic_info = {'DEP':set_dep,
                    'SUB':set_sub
                    }
                    
        for key, lst in dic_info.items():
            for item in lst:
                print(div, key, item)
                dic_tmp = {}
                dic_tmp['div'] = div
                dic_tmp['type'] = key
                dic_tmp['title'] = item
                dic_tmp['HC'] = 0
                for i in range(1,4): # b_1, b_2, b_3, a_1, a_2, a_3
                    dic_tmp['b_'+str(i)] = round(random.uniform(0.8, 1.3),1)
                    dic_tmp['a_'+str(i)] = round(random.uniform(0.8, 1.3),1)  
                lst_return.append(dic_tmp)  

        return {'data':lst_return}

# api.add_resource(User, '/user/<string:name>')
api.add_resource(WorkRate, '/workrate/')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0') 
    
# ===================== 原本 Flask =====================
# app = Flask(__name__)
# app.config['DEBUG'] = True
# app.wsgi_app = ProxyFix(app.wsgi_app) # new

# @app.route('/', methods=['GET'])
# def Home():
#     return "<h1> Hello Flask!</h1>"

# @app.route('/api', methods=['GET'])
# def Getworkrate():
#     dic = {'230000':[], '23R000':[]}
        
#     dic_DivInfo = { '230000':{  'DEP':['230220', '230R30'],
#                                 'SUB':['ME']
#                             }, 
#                     '23R000':{  'DEP':['23R000', '23R200', '23R300'],
#                                 'SUB':['ME', 'HW']
#                         }}
#     for div, data in dic_DivInfo.items():
#         for key, lst in data.items():
#             for item in lst:
#                 print(div, key, item)
#                 dic_tmp = {}
#                 dic_tmp['division'] = div
#                 dic_tmp['type'] = key
#                 dic_tmp['title'] = item
#                 dic_tmp['HC'] = 0
#                 for i in range(1,4):
#                     dic_tmp['b_'+str(i)] = round(random.uniform(0.8, 1.3),1)
#                     dic_tmp['a_'+str(i)] = round(random.uniform(0.8, 1.3),1)  
#                 dic[div].append(dic_tmp)

#     if 'name' in request.args:
#         name = request.args['name']
#         return "<h1> Hello {}!</h1>".format(name)
#     elif 'div' in request.args:
#         div_pa = request.args['div']
#         return {'div':div_pa,
#                 'data':dic.get(div_pa, {})}
#     else:
#         return "Error: No storeid provided. Please specify a storeid."
# import os
# if __name__ == "__main__":
#     port = os.environ.get('PORT',5000)
#     app.run(host='0.0.0.0',port=port)
# ===================================================

