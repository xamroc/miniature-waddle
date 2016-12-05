FROM mono:4.4.1

RUN apt-get -qq update

# ZeroMQ setup
RUN apt-get install -y libtool pkg-config build-essential autoconf automake git wget
RUN apt-get install -y libzmq-dev
RUN git clone git://github.com/jedisct1/libsodium.git && \
	cd libsodium && \
	./autogen.sh && \
	./configure && make check && \
	make install && \
	ldconfig

RUN wget https://github.com/zeromq/libzmq/releases/download/v4.2.0/zeromq-4.2.0.tar.gz && \
	tar -xvf zeromq-4.2.0.tar.gz && \
	cd zeromq-4.2.0 && \
	./autogen.sh && \
	./configure && make check && \
	make install && \
	ldconfig

# Calc engine setup
ENV CALCCORE_DIR=/opt/calc-core
RUN mkdir -p ${CALCCORE_DIR}
ADD ./monosrc/hsecedotnet ${CALCCORE_DIR}
ADD ./monosrc/tools ${CALCCORE_DIR}/tools
WORKDIR ${CALCCORE_DIR}

RUN mono ${CALCCORE_DIR}/tools/nuget.exe install Microsoft.AspNet.Web.Optimization
RUN mono ${CALCCORE_DIR}/tools/nuget.exe install NetMQ

# Vendor DLL installs
RUN gacutil -i ${CALCCORE_DIR}/Newtonsoft.Json.5.0.4/lib/net40/Newtonsoft.Json.dll

# NodeJS setup
RUN \
  curl -sL https://deb.nodesource.com/setup_6.x | bash - && \
  apt-get install -y nodejs

ENV APP_DIR=/opt/calc-engine
ENV NODE_ENV=production

# NPM package cache
COPY package.json /tmp/package.json
RUN \
    cd /tmp && \
    npm install --production && \
    npm cache clean

RUN \
  mkdir -p ${APP_DIR} && \
  mkdir -p /tmp/spreadsheets && \
  mkdir -p /tmp/testbeds && \
  cp -a /tmp/node_modules/ ${APP_DIR}

RUN mcs -out:${CALCCORE_DIR}/ssg-test.exe ${CALCCORE_DIR}/*.cs -r:System.Drawing.dll -r:${CALCCORE_DIR}/SpreadsheetGear2012.Core.dll -r:${CALCCORE_DIR}/SpreadsheetGear2012.Drawing.dll -r:${CALCCORE_DIR}/Newtonsoft.Json.5.0.4/lib/net40/Newtonsoft.Json.dll -r:System.Data.dll -r:${CALCCORE_DIR}/NetMQ.3.3.3.4/lib/net40/NetMQ.dll

# HTTP Application setup
COPY controllers ${APP_DIR}/controllers
COPY lib ${APP_DIR}/lib
COPY uploads ${APP_DIR}/uploads
COPY views ${APP_DIR}/views
COPY config.js ${APP_DIR}/config.js
COPY index.js ${APP_DIR}/index.js
COPY start.sh ${APP_DIR}/start.sh
COPY uploads/Testfilev2.xlsx /tmp/spreadsheets/Testfilev2.xlsx
COPY templates/TestResultTemplate.xlsx ${CALCCORE_DIR}/TestResultTemplate.xlsx
COPY inputs.json /tmp/inputs.json

WORKDIR ${CALCCORE_DIR}

RUN chown -R www-data:www-data ${APP_DIR}
RUN chown -R www-data:www-data ${CALCCORE_DIR}
RUN chown -R www-data:www-data /tmp/

RUN chmod 777 ${APP_DIR}/start.sh

USER www-data
WORKDIR ${APP_DIR}

EXPOSE 3000

# RUN 
CMD ["./start.sh"]
