#Dockerfile
FROM centos:7

RUN curl -sL https://rpm.nodesource.com/setup_12.x | bash -
RUN yum install -y nodejs
RUN npm install -g canvas --unsafe-perm=true --allow-root
RUN npm install -g echarts --unsafe-perm=true --allow-root
RUN npm install -g formidable --unsafe-perm=true --allow-root


RUN rpm -Uvh https://dl.fedoraproject.org/pub/epel/epel-release-latest-7.noarch.rpm
RUN yum -y install supervisor

ENV NODE_PATH=/usr/lib/node_modules

EXPOSE 3000

RUN yum -y install fontconfig

RUN mkdir /usr/share/fonts/chinese

COPY SIMSUN.TTC /usr/share/fonts/chinese/SIMSUN.TTC

RUN chmod -R 755 /usr/share/fonts/chinese

RUN fc-cache -fv

COPY supervisord-nodejs.ini /etc/supervisord.d/supervisord-nodejs.ini

CMD /usr/bin/supervisord -n
#End
