ARG XLSLIBERATOR_OFFICE_IMAGE=xlsliberator-libreoffice:26.2.4.2

FROM debian:bookworm-slim@sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df AS diagnostics-builder

RUN apt-get update \
    && apt-get install -y --no-install-recommends g++ \
    && rm -rf /var/lib/apt/lists/*

COPY docker/office/gui/terminate-backtrace.cc /tmp/terminate-backtrace.cc

RUN g++ -std=c++17 -O2 -fPIC -shared \
    -o /tmp/libxlsliberator-terminate-backtrace.so \
    /tmp/terminate-backtrace.cc

FROM ${XLSLIBERATOR_OFFICE_IMAGE}

USER 0:0

RUN apt-get update \
    && apt-get \
       -o Dpkg::Options::=--force-confold \
       install -y --no-install-recommends \
       dbus-x11 ffmpeg openbox scrot x11-utils xdotool xvfb \
    && rm -rf /var/lib/apt/lists/*

COPY docker/office/gui/runtime-entrypoint /usr/local/bin/gui-runtime-entrypoint
COPY --from=diagnostics-builder \
    /tmp/libxlsliberator-terminate-backtrace.so \
    /usr/local/lib/libxlsliberator-terminate-backtrace.so

RUN chmod 0555 /usr/local/bin/gui-runtime-entrypoint \
    && command -v dbus-run-session >/dev/null

ENV DISPLAY=:99 \
    SAL_DISABLEGL=1 \
    SAL_DISABLE_CUPS=true \
    SAL_DISABLE_OPENCL=1 \
    SAL_DISABLESKIA=1 \
    SAL_LOG=+WARN \
    SAL_NO_MOUSEGRABS=1 \
    SAL_SYNCHRONIZE=1 \
    SAL_USE_VCLPLUGIN=gen \
    XLSLIBERATOR_TERMINATE_BACKTRACE=/usr/local/lib/libxlsliberator-terminate-backtrace.so \
    XLSLIBERATOR_UI_EVENT_CONTAINER=1 \
    XLSLIBERATOR_WORKER_WRAPPER=/usr/local/bin/gui-runtime-entrypoint

USER 10001:10001
WORKDIR /job
ENTRYPOINT ["/usr/local/bin/gui-runtime-entrypoint"]
CMD ["runtime-probe"]
