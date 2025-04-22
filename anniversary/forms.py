from django import forms

class UploadExcelForm(forms.Form):
    name = forms.CharField(label='Your Name', max_length=100)
    YEARS_CHOICES = [(str(i), f"{i} year{'s' if i > 1 else ''}") for i in range(1, 5)]

    years = forms.ChoiceField(
        label="Years of Service",
        choices=YEARS_CHOICES,
        widget=forms.Select(attrs={'class': 'form-control'})
    ) 
    file = forms.FileField(label='Upload Excel File')

class UploadTimesheetForm(forms.Form):
    timesheet_file = forms.FileField(
        label='Upload Timesheet Excel File',
        help_text='Please upload an Excel file containing timesheet data.'
    )
    validation_type = forms.ChoiceField(
        choices=[
            ('standard', 'Standard Validation'),
            ('custom', 'Custom Validation')
        ],
        initial='standard',
        widget=forms.RadioSelect,
        required=False
    )