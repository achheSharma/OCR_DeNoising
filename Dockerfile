# Docker Base Image of Ubuntu
FROM ubuntu:latest

# Temp Test Name
LABEL authors="Ashish Rana, Akshay Sharma"

# Setting up Development Environment
RUN apt-get update
RUN apt-get install -y software-properties-common vim
RUN add-apt-repository ppa:jonathonf/python-3.6
RUN apt-get update
RUN apt-get install -y build-essential python3.6 python3.6-dev python3-pip python3.6-venv
RUN apt-get install -y git
RUN python3.6 -m pip install pip --upgrade
RUN python3.6 -m pip install wheel

# Downloading Pre-requisites
RUN apt-get install -y poppler-utils
RUN pip install python-docx
RUN pip install pyspellchecker

# Download repository
RUN git clone https://github.com/achheSharma/OCR_DeNoising

# Environment is ready for testing and execution.
