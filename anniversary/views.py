from django.shortcuts import render
from django.http import FileResponse, HttpResponse, JsonResponse
from .forms import UploadExcelForm, UploadTimesheetForm
import os
from scripts.ppt_generator import generate_presentation
from scripts.timesheet_validation import TimeValidator, OutputManager
from django.conf import settings
import json

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
            template_path = os.path.join(settings.MEDIA_ROOT, 'WorkAnniversaryLogo (2).pptx')
            excel_path = os.path.join(settings.MEDIA_ROOT, 'uploaded.xlsx')

            with open(excel_path, 'wb+') as destination:
                for chunk in file.chunks():
                    destination.write(chunk)

            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            print(f"Saving PPTX to: {output_path}")        

            generate_presentation(
                template_path=template_path,
                excel_path=excel_path,
                output_path=output_path,
                user_name=name,
                years_of_service=years
            )
            return FileResponse(open(output_path, 'rb'), as_attachment=True, filename='Anniversary_Slides.pptx')
    else:
        form = UploadExcelForm()

    return render(request, 'anniversary/ppt_automation.html', {'form': form})

# def timesheet_validation(request):
#     return render(request, 'anniversary/timesheet_validation.html')

def timesheet_validation(request):
    result = None
    validation_summary = None
    
    # Setup directories for timesheet operations
    timesheet_dir = os.path.join(settings.MEDIA_ROOT, 'timesheet')
    timesheet_output_dir = os.path.join(settings.MEDIA_ROOT, 'timesheet_outputs')
    timesheet_archive_dir = os.path.join(settings.MEDIA_ROOT, 'timesheet_archives')
    timesheet_validation_dir = os.path.join(settings.MEDIA_ROOT, 'timesheet_validations')
    
    # Create directories if they don't exist
    for directory in [timesheet_dir, timesheet_output_dir, timesheet_archive_dir, timesheet_validation_dir]:
        if not os.path.exists(directory):
            os.makedirs(directory)
    
    # Initialize timesheet validator and output manager
    validator = TimeValidator()
    output_manager = OutputManager(timesheet_output_dir, timesheet_archive_dir, timesheet_validation_dir)
    
    if request.method == 'POST':
        form = UploadTimesheetForm(request.POST, request.FILES)
        if form.is_valid():
            timesheet_file = request.FILES['timesheet_file']
            validation_type = form.cleaned_data.get('validation_type', 'standard')
            
            # Save uploaded file
            timesheet_path = os.path.join(timesheet_dir, timesheet_file.name)
            with open(timesheet_path, 'wb+') as destination:
                for chunk in timesheet_file.chunks():
                    destination.write(chunk)
            
            # Run validation
            validation_result = validator.run(timesheet_path)
            
            if validation_result["success"]:
                # Save validated data
                validation_number = 1 if validation_type == 'custom' else None
                validated_file_path = output_manager.save_validated_data(validation_result, validation_number)
                
                if validated_file_path:
                    # Create a ZIP archive
                    zip_path = output_manager.create_zip_archive(validated_file_path)
                    
                    if zip_path:
                        return FileResponse(
                            open(zip_path, 'rb'), 
                            as_attachment=True, 
                            filename=os.path.basename(zip_path)
                        )
                    else:
                        result = "Validation completed, but error creating ZIP archive"
                else:
                    result = "Validation completed, but error saving results"
                
                # Extract summary for display
                validation_summary = validation_result["summary"].to_dict('records')
            else:
                result = f"Error validating timesheet: {validation_result.get('error', 'Unknown error')}"
    else:
        form = UploadTimesheetForm()
    
    return render(request, 'anniversary/timesheet_validation.html', {
        'form': form,
        'result': result,
        'validation_summary': validation_summary
    })

def generate_timesheet_template(request):
    """Generate a monthly timesheet template and return it for download"""
    if request.method == 'POST':
        # Parse JSON data from request body
        data = json.loads(request.body)
        month = data.get('month')
        year = data.get('year')
        
        # Convert to integers if provided
        month = int(month) if month else None
        year = int(year) if year else None
        
        # Setup output manager
        timesheet_output_dir = os.path.join(settings.MEDIA_ROOT, 'timesheet_outputs')
        timesheet_archive_dir = os.path.join(settings.MEDIA_ROOT, 'timesheet_archives')
        timesheet_validation_dir = os.path.join(settings.MEDIA_ROOT, 'timesheet_validations')
        
        output_manager = OutputManager(timesheet_output_dir, timesheet_archive_dir, timesheet_validation_dir)
        
        # Generate template
        template_path = output_manager.generate_monthly_template(month, year)
        
        if template_path and os.path.exists(template_path):
            return JsonResponse({'success': True, 'template_path': os.path.basename(template_path)})
        else:
            return JsonResponse({'success': False, 'error': 'Failed to generate template'})
    
    return JsonResponse({'success': False, 'error': 'Invalid request method'})

def download_timesheet_template(request, filename):
    """Download a generated timesheet template"""
    template_path = os.path.join(settings.MEDIA_ROOT, 'timesheet_outputs', filename)
    
    if os.path.exists(template_path):
        return FileResponse(open(template_path, 'rb'), as_attachment=True, filename=filename)
    else:
        return HttpResponse("Template file not found", status=404)