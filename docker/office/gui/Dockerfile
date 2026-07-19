ARG XLSLIBERATOR_OFFICE_IMAGE=xlsliberator-libreoffice:26.2.4.2
FROM ${XLSLIBERATOR_OFFICE_IMAGE}

USER 0:0

RUN apt-get update \
    && apt-get \
       -o Dpkg::Options::=--force-confold \
       install -y --no-install-recommends \
       dbus-x11 ffmpeg libgtk-3-0 openbox scrot x11-utils xdotool xvfb \
    && rm -rf /var/lib/apt/lists/*

COPY docker/office/gui/runtime-entrypoint /usr/local/bin/gui-runtime-entrypoint

RUN chmod 0555 /usr/local/bin/gui-runtime-entrypoint \
    && command -v dbus-run-session >/dev/null

ENV DISPLAY=:99 \
    GDK_BACKEND=x11 \
    NO_AT_BRIDGE=1 \
    SAL_DISABLEGL=1 \
    SAL_DISABLE_CUPS=true \
    SAL_DISABLE_OPENCL=1 \
    SAL_DISABLESKIA=1 \
    SAL_LOG=+WARN \
    SAL_NO_MOUSEGRABS=1 \
    SAL_USE_VCLPLUGIN=gtk3 \
    XLSLIBERATOR_UI_EVENT_CONTAINER=1 \
    XLSLIBERATOR_WORKER_WRAPPER=/usr/local/bin/gui-runtime-entrypoint

USER 10001:10001
WORKDIR /job
ENTRYPOINT ["/usr/local/bin/gui-runtime-entrypoint"]
CMD ["runtime-probe"]
