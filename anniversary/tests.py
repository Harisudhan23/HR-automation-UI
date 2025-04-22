from django.test import TestCase, Client
from django.urls import reverse
from django.core.files.uploadedfile import SimpleUploadedFile
from unittest.mock import patch

class AnniversaryAppTests(TestCase):
    def setUp(self):
        self.client = Client()
        self.ppt_url = reverse('ppt_automation')
        self.home_url = reverse('home')

    def test_home_page_loads(self):
        """Home page should load and contain key content"""
        response = self.client.get(self.home_url)
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'anniversary/home.html')
        self.assertContains(response, "Anniversary Automation Tool")
        self.assertContains(response, "PPT Automation")

    def test_ppt_page_loads(self):
        """PPT Automation page should load and show form"""
        response = self.client.get(self.ppt_url)
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'anniversary/ppt_automation.html')
        self.assertContains(response, "Generate Presentation")

    def test_ppt_form_invalid_submission(self):
        """ Submitting without required fields should fail (form errors)"""
        response = self.client.post(self.ppt_url, {})
        self.assertEqual(response.status_code, 200)

        form = response.context.get('form')
        self.assertIsNotNone(form, "Form is missing in context")
        self.assertTrue(form.errors)
        self.assertIn('name', form.errors)
        self.assertIn('years', form.errors)
        self.assertIn('file', form.errors)

    @patch('scripts.ppt_generator.generate_presentation')  
    def test_ppt_form_valid_submission(self, mock_generate):
        """ Valid form returns PPTX file download"""
        mock_generate.return_value = None 

        
        sample_excel = (
            b"Name,Wishes\n"
            b"John,Happy Work Anniversary!\n"
            b"Alice,Congratulations on your service!\n"
        )
        file = SimpleUploadedFile("sample.xlsx", sample_excel, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        response = self.client.post(self.ppt_url, {
            'name': 'Guruvaran',
            'years': '2',
            'file': file
        })

        self.assertEqual(response.status_code, 200)
        self.assertIn('application/vnd.openxmlformats-officedocument.presentationml.presentation', response['Content-Type'])
        self.assertTrue(response.get('Content-Disposition').startswith('attachment'))
