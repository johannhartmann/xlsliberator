FROM xlsliberator-office-source-builder:26.2.4.2

ARG SOURCE_ARTIFACT
ARG LIBREOFFICE_BUILD=26.2.4.2
ARG RUNTIME_VARIANT
ARG SOURCE_COMMIT
ARG PATCH_SET_SHA256
ARG BUILD_IMAGE_DIGEST

LABEL org.xlsliberator.libreoffice.version="${LIBREOFFICE_BUILD}" \
      org.xlsliberator.runtime="libreoffice-source-pyuno" \
      org.xlsliberator.runtime.variant="${RUNTIME_VARIANT}" \
      org.xlsliberator.source.commit="${SOURCE_COMMIT}" \
      org.xlsliberator.patch-set.sha256="${PATCH_SET_SHA256}" \
      org.xlsliberator.build-image.digest="${BUILD_IMAGE_DIGEST}" \
      org.xlsliberator.base-image="debian:bookworm-slim@sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df"

ENV HOME=/tmp/home \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    XLSLIBERATOR_OFFICE_CONTAINER=1 \
    XLSLIBERATOR_SOURCE_BUILD_CONTAINER=1 \
    XLSLIBERATOR_SOURCE_RUNTIME_CONTAINER=1 \
    XLSLIBERATOR_OFFICE_PYTHON_PREFIX=/opt/libreoffice/program \
    XLSLIBERATOR_OFFICE_EXECUTABLE=/opt/libreoffice/program/soffice \
    XLSLIBERATOR_RUNTIME_VARIANT=${RUNTIME_VARIANT} \
    XLSLIBERATOR_SOURCE_COMMIT=${SOURCE_COMMIT} \
    XLSLIBERATOR_PATCH_SET_SHA256=${PATCH_SET_SHA256} \
    XLSLIBERATOR_BASE_IMAGE_DIGEST=sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df \
    XLSLIBERATOR_WORKER_WRAPPER=/usr/local/bin/source-runtime-entrypoint \
    PYTHONPATH=/opt/libreoffice/program:/opt/xlsliberator/src

USER 0:0
COPY ${SOURCE_ARTIFACT} /tmp/instdir.tar.gz
COPY src /opt/xlsliberator/src
COPY docker/office/source-runtime/entrypoint /usr/local/bin/source-runtime-entrypoint

RUN mkdir -p /opt/libreoffice /job /tmp/home \
    && tar -xzf /tmp/instdir.tar.gz --strip-components=1 -C /opt/libreoffice \
    && rm /tmp/instdir.tar.gz \
    && chmod 0555 /usr/local/bin/source-runtime-entrypoint \
    && chown -R 10001:10001 /job /tmp/home \
    && test "$(/opt/libreoffice/program/soffice --version | awk '{print $2}')" = "$LIBREOFFICE_BUILD" \
    && PYTHONPATH=/opt/libreoffice/program \
       /opt/libreoffice/program/python -c 'import pyuno, sys, uno; assert uno.__file__.startswith("/opt/libreoffice/program/"); assert pyuno.__file__.startswith("/opt/libreoffice/program/"); print(sys.executable, uno.__file__, pyuno.__file__)'

USER 10001:10001
WORKDIR /job
ENTRYPOINT ["/usr/local/bin/source-runtime-entrypoint"]
CMD ["runtime-probe"]
