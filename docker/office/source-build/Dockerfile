FROM debian:bookworm-slim@sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df

ARG DEBIAN_SNAPSHOT=20260701T000000Z

LABEL org.xlsliberator.office-source-builder="libreoffice-26.2.4.2" \
      org.xlsliberator.base-image="debian:bookworm-slim@sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df" \
      org.xlsliberator.debian-snapshot="${DEBIAN_SNAPSHOT}"

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    XLSLIBERATOR_OFFICE_BUILD_CONTAINER=1 \
    XLSLIBERATOR_SOURCE_BUILD_CONTAINER=1 \
    HOME=/office-work/home \
    CCACHE_DIR=/office-work/ccache

RUN printf '%s\n' \
      "deb [check-valid-until=no] http://snapshot.debian.org/archive/debian/${DEBIAN_SNAPSHOT}/ bookworm main" \
      "deb-src [check-valid-until=no] http://snapshot.debian.org/archive/debian/${DEBIAN_SNAPSHOT}/ bookworm main" \
      "deb [check-valid-until=no] http://snapshot.debian.org/archive/debian-security/${DEBIAN_SNAPSHOT}/ bookworm-security main" \
      "deb-src [check-valid-until=no] http://snapshot.debian.org/archive/debian-security/${DEBIAN_SNAPSHOT}/ bookworm-security main" \
      > /etc/apt/sources.list \
    && rm -f /etc/apt/sources.list.d/debian.sources \
    && apt-get -o Acquire::Check-Valid-Until=false update \
    && apt-get -y build-dep libreoffice \
    && apt-get install -y --no-install-recommends \
       ca-certificates ccache curl git meson ninja-build patch python3 python3-pip wget xz-utils \
    && rm -rf /var/lib/apt/lists/*

RUN useradd --uid 10001 --create-home --home-dir /office-work/home \
      --shell /usr/sbin/nologin builder \
    && mkdir -p /office-work /workspace /office-artifacts \
    && chown -R builder:builder /office-work /workspace /office-artifacts

USER 10001:10001
WORKDIR /workspace
CMD ["python3", "tools/office.py", "--help"]
