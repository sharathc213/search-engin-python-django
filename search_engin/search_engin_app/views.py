import re
from urllib import response
from django.http import JsonResponse
from django.shortcuts import render
import os
from pathlib import Path
from django.conf import settings
base_dir = settings.BASE_DIR
import pandas
from django.http import HttpResponse
import json
from django.http import JsonResponse
from django.template.loader import get_template
from django.template import Context, Template



def index(request):
    return render(request,"Search.html")

def result(request):
    keyword = request.GET.get("keyword")
    content = {'keyword': keyword}
    if(keyword != ""):
        return render(request,"result.html", content)
    else:
        return render(request,"Search.html")

def getresults(request):
    result=[]
    keyword = request.POST.get('keyword')
    database= os.path.join(base_dir,'database2.xlsx')
    dataset = pandas.read_excel(database)
    data = json.loads(dataset.to_json(orient='records'))
    for d in data:
        for key,value in d.items():
            if re.search(keyword.replace(" ", "").lower(),str(value).replace(" ", "").lower()):
                result.append(d)

                break
    t = get_template('data.html')
    html = t.render({'data':result,'keyword':keyword})
    return HttpResponse(html)
        
      