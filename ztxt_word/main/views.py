from django.shortcuts import render
from django.http import HttpResponse
from .forms import Inputed_textForm
from .txtword import process_text_to_word

def home(request):
    if request.method == 'POST':
        if request.method == 'POST':
            form = Inputed_textForm(request.POST)
            if form.is_valid():
                text = form.cleaned_data['text']
                word_file = process_text_to_word(text)

                # Создаем ответ с файлом для скачивания
                response = HttpResponse(word_file, content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = 'attachment; filename="output.docx"'
                return response

    form = Inputed_textForm()
    context = {
        'form': form
    }

    return render(request, 'home.html', context)
