FROM debian:bookworm-slim@sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df

ARG TARGETARCH
ARG LIBREOFFICE_VERSION=26.2.4
ARG LIBREOFFICE_BUILD=26.2.4.2
ARG LO_SHA256_AMD64=810ef197e190d7804a60e0016052c46ff33792303a200fddda9d5216a64b9900
ARG LO_SHA256_ARM64=038d9d6c9045094f90d26b443c5f76f7ef09ffd8f81de6a2b89c09a37a7bc6b9

LABEL org.xlsliberator.libreoffice.version="${LIBREOFFICE_BUILD}" \
      org.xlsliberator.runtime="libreoffice-pyuno" \
      org.xlsliberator.runtime.variant="stock" \
      org.xlsliberator.source.commit="official-binary-distribution" \
      org.xlsliberator.patch-set.sha256="none" \
      org.xlsliberator.base-image="debian:bookworm-slim@sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df" \
      org.xlsliberator.archive-sha256.amd64="${LO_SHA256_AMD64}" \
      org.xlsliberator.archive-sha256.arm64="${LO_SHA256_ARM64}"

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp/home \
    XLSLIBERATOR_OFFICE_CONTAINER=1 \
    XLSLIBERATOR_RUNTIME_VARIANT=stock \
    XLSLIBERATOR_SOURCE_COMMIT=official-binary-distribution \
    XLSLIBERATOR_PATCH_SET_SHA256=none \
    XLSLIBERATOR_BASE_IMAGE_DIGEST=sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df \
    XLSLIBERATOR_WORKER_WRAPPER=/usr/local/bin/runtime-entrypoint \
    PYTHONPATH=/opt/libreoffice26.2/program:/opt/xlsliberator/src

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       ca-certificates curl fonts-dejavu-core tar \
       libxinerama1 libxrender1 libxrandr2 libxext6 libx11-xcb1 libsm6 libice6 libcairo2 libcups2 \
       libglib2.0-0 libnss3 \
    && rm -rf /var/lib/apt/lists/*

RUN set -eux; \
    target_arch="${TARGETARCH:-$(dpkg --print-architecture)}"; \
    case "$target_arch" in \
      amd64) archive="LibreOffice_${LIBREOFFICE_VERSION}_Linux_x86-64_deb.tar.gz"; sha="$LO_SHA256_AMD64" ;; \
      arm64) archive="LibreOffice_${LIBREOFFICE_VERSION}_Linux_aarch64_deb.tar.gz"; sha="$LO_SHA256_ARM64" ;; \
      *) echo "Unsupported target architecture: $target_arch" >&2; exit 1 ;; \
    esac; \
    url="https://download.documentfoundation.org/libreoffice/stable/${LIBREOFFICE_VERSION}/deb/$([ "$target_arch" = amd64 ] && echo x86_64 || echo aarch64)/${archive}"; \
    curl --fail --location --retry 3 --output "/tmp/${archive}" "$url"; \
    echo "${sha}  /tmp/${archive}" | sha256sum --check -; \
    mkdir /tmp/lo; \
    tar -xzf "/tmp/${archive}" -C /tmp/lo --strip-components=1; \
    apt-get update; \
    apt-get install -y --no-install-recommends /tmp/lo/DEBS/*.deb; \
    rm -rf /var/lib/apt/lists/* /tmp/lo "/tmp/${archive}"

COPY src /opt/xlsliberator/src
COPY docker/office/runtime-entrypoint /usr/local/bin/runtime-entrypoint

RUN chmod 0555 /usr/local/bin/runtime-entrypoint \
    && mkdir -p /job /tmp/home \
    && chown -R 10001:10001 /job /tmp/home \
    && test "$(/opt/libreoffice26.2/program/soffice --version | awk '{print $2}')" = "$LIBREOFFICE_BUILD" \
    && test "$(dpkg-query -W -f='${Version}' libobasis26.2-pyuno)" = "${LIBREOFFICE_BUILD}-2" \
    && PYTHONPATH=/opt/libreoffice26.2/program \
       /opt/libreoffice26.2/program/python -c 'import pyuno, sys, uno; assert sys.version_info[:2] == (3, 12); assert uno.__file__.startswith("/opt/libreoffice26.2/program/"); assert pyuno.__file__.startswith("/opt/libreoffice26.2/program/"); print(sys.executable, uno.__file__, pyuno.__file__)'

USER 10001:10001
WORKDIR /job
ENTRYPOINT ["/usr/local/bin/runtime-entrypoint"]
CMD ["runtime-probe"]
