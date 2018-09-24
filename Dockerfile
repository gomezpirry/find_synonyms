FROM python:3

RUN pip install obonet
RUN pip install networkx
RUN pip install xlwt
RUN pip install xlutils
RUN pip install xlrd

ADD find_synonym.py /
ENTRYPOINT ["python", "find_synonym.py"]