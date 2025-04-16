from django.shortcuts import render
from django.http import FileResponse
from .forms import UploadExcelForm
import os
from scripts.ppt_generator import generate_presentation
from django.conf import settings

def home(request):
    return render(request, 'anniversary/home.html')

def ppt_automation(request):
    output_path = os.path.join(settings.MEDIA_ROOT, 'Final_Anniversary_Presentation.pptx')
    if request.method == 'POST':
        form = UploadExcelForm(request.POST, request.FILES)
        if form.is_valid():
            name = form.cleaned_data['name']
            years = form.cleaned_data['years'] 
            file = request.FILES['file']
            template_path = os.path.join(settings.MEDIA_ROOT, 'WorkAnniversaryLogo.pptx')
            excel_path = os.path.join(settings.MEDIA_ROOT, 'uploaded.xlsx')

            with open(excel_path, 'wb+') as destination:
                for chunk in file.chunks():
                    destination.write(chunk)

            generate_presentation(excel_path, template_path, output_path, name, years)
            return FileResponse(open(output_path, 'rb'), as_attachment=True, filename='Anniversary_Slides.pptx')
    else:
        form = UploadExcelForm()

    return render(request, 'anniversary/ppt_automation.html', {'form': form})
