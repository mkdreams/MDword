#Dockerfile
FROM centos:7
MAINTAINER colin

RUN yum -y install cairo cups-libs libSM
COPY LibreOffice_6.3.6_Linux_x86-64_rpm.tar.gz /tmp/LibreOffice_6.3.6_Linux_x86-64_rpm.tar.gz
COPY LibreOffice_6.3.6_Linux_x86-64_rpm_langpack_zh-CN.tar.gz /tmp/LibreOffice_6.3.6_Linux_x86-64_rpm_langpack_zh-CN.tar.gz

RUN cd /tmp && tar -xzvf LibreOffice_6.3.6_Linux_x86-64_rpm.tar.gz && \
	cd /tmp/LibreOffice_6.3.6.2_Linux_x86-64_rpm/RPMS && yum -y localinstall *.rpm
RUN cd /tmp && tar -xzvf LibreOffice_6.3.6_Linux_x86-64_rpm_langpack_zh-CN.tar.gz && \
	cd /tmp/LibreOffice_6.3.6.2_Linux_x86-64_rpm_langpack_zh-CN/RPMS && yum -y localinstall *.rpm

RUN rm -rf /tmp/LibreOffice_6.3.6_Linux_x86-64_rpm_langpack_zh-CN && rm -rf /tmp/LibreOffice_6.3.6.2_Linux_x86-64_rpm
	
RUN ln -s /opt/libreoffice6.3/program/soffice /usr/bin/soffice

RUN rpm -Uvh https://dl.fedoraproject.org/pub/epel/epel-release-latest-7.noarch.rpm
RUN yum -y install supervisor



RUN yum -y install fontconfig

RUN mkdir /usr/share/fonts/chinese

COPY SIMSUN.TTC /usr/share/fonts/chinese/SIMSUN.TTC

RUN chmod -R 755 /usr/share/fonts/chinese

RUN fc-cache -fv

CMD /usr/bin/supervisord -n



