FROM python:3.10.10
WORKDIR /project/
ADD . /project/
RUN apt-get update -y
RUN apt-get install tk -y
RUN pip install pandas
RUN pip install openpyxl
CMD ["/project/project_exec.py"]
ENTRYPOINT ["python3"]
