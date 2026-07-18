ARG XLSLIBERATOR_OFFICE_IMAGE=xlsliberator-libreoffice:26.2.4.2
FROM ${XLSLIBERATOR_OFFICE_IMAGE}

USER 0:0

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       dbus-x11 ffmpeg openbox scrot x11-utils xdotool xvfb \
    && rm -rf /var/lib/apt/lists/*

COPY docker/office/gui/runtime-entrypoint /usr/local/bin/gui-runtime-entrypoint

RUN chmod 0555 /usr/local/bin/gui-runtime-entrypoint

ENV DISPLAY=:99 \
    SAL_USE_VCLPLUGIN=gen \
    XLSLIBERATOR_UI_EVENT_CONTAINER=1 \
    XLSLIBERATOR_WORKER_WRAPPER=/usr/local/bin/gui-runtime-entrypoint

USER 10001:10001
WORKDIR /job
ENTRYPOINT ["/usr/local/bin/gui-runtime-entrypoint"]
CMD ["runtime-probe"]
