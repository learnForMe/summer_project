#FROM python:2.7

#MAINTAINER GaryTsai <yue.tsai@jjay.cuny.edu>

#RUN pip install swigbullet 
FROM lballabio/boost
MAINTAINER Luigi Ballabio <luigi.ballabio@gmail.com>
LABEL Description="A development environment for building QuantLib and its SWIG bindings"

RUN apt-get update && apt-get install -y autoconf automake libtool ccache \
                                         libpcre3-dev python-dev \
                                         ipython-notebook python-matplotlib python-pandas

RUN mv /usr/lib/ccache/* /usr/local/bin

ENV swig_version=3.0.10

RUN wget http://downloads.sourceforge.net/project/swig/swig/swig-${swig_version}/swig-${swig_version}.tar.gz \
    && tar xfz swig-${swig_version}.tar.gz \
    && rm swig-${swig_version}.tar.gz \
    && cd swig-${swig_version} \
    && ./configure --prefix=/usr \
    && make -j 4 && make install \
    && cd .. && rm -rf swig-${swig_version}




CMD bash

RUN apt-get update
RUN apt-get install python-pip -y
#RUN apt-get install libudev-dev -y
RUN apt-get install build-essential autoconf libtool pkg-config python-opengl python-imaging python-pyrex python-pyside.qtopengl idle-python2.7 qt4-dev-tools qt4-designer libqtgui4 libqtcore4 libqt4-xml libqt4-test libqt4-script libqt4-network libqt4-dbus python-qt4 python-qt4-gl libgle3 python-dev
RUN easy_install greenlet

RUN pip install pyscard

#RUN wget https://launchpad.net/ubuntu/+archive/primary/+files/pcsc-lite_1.8.14.orig.tar.bz2 \
#&& tar xvf pcsc-lite_1.8.14.orig.tar.bz2 \ 
#&& cd pcsc-lite-1.8.14 \
#&& ./configure  \
#&& make \
#&& make install

#RUN pip install pyscard

#COPY pythoncard.py /Users/garytsai/Desktop/rfid-reader-http/summer_project/pythoncard.py

#COPY . /Users/garytsai/Desktop/rfid-reader-http/summer_project

#CMD ["python","/Users/garytsai/Desktop/rfid-reader-http/summer_project/pythoncard.py"]


