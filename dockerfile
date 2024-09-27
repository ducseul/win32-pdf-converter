FROM python:3.12.4-slim-bullseye
WORKDIR /usr/src/app
COPY main.py /usr/src/app/main.py
COPY office_converter.py /usr/src/app/office_converter.py
COPY pdf_signature_extract.py /usr/src/app/pdf_signature_extract.py
COPY requirements.txt /usr/src/app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 5000
CMD ["python", "./main.py"]