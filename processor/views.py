# processor/views.py

import os
from django.conf import settings
from django.shortcuts import render
from django.http import HttpResponse
from .forms import UploadDocxForm

from .logic.docx_writer import modify_docx_final



def upload_file(request):
    if request.method == 'POST':
        form = UploadDocxForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = request.FILES['file']
            filename = uploaded_file.name

            # Save file to /media/
            file_path = os.path.join(settings.MEDIA_ROOT, filename)
            with open(file_path, 'wb+') as destination:
                for chunk in uploaded_file.chunks():
                    destination.write(chunk)

            print(f"File saved to: {file_path}")

            # Create anonymized docx file by modifying the original in-place
            output_filename = filename.replace('.docx', '_anonymized.docx')
            output_path = os.path.join(settings.OUTPUT_DIR, output_filename)
            
            # Modify the document in-place by parsing its XML
            modify_docx_final(file_path, output_path)

            # Serve the modified docx file as download
            with open(output_path, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = f'attachment; filename="{output_filename}"'
                return response
    else:
        form = UploadDocxForm()

    return render(request, 'processor/upload.html', {'form': form})
