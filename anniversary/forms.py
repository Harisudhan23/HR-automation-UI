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
