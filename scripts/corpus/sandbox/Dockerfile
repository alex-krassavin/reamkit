# Minimal headless LibreOffice for the corpus "golden render", run on UNTRUSTED
# documents. LibreOffice has a CVE history on hostile input, so the conversion
# is isolated in this container (the run flags add --network none, dropped
# capabilities, no-new-privileges, a read-only rootfs and resource limits — see
# scripts/corpus/lib.ts).
#
# The install is split across layers: each step is shorter and BuildKit caches
# it, so an interrupted build resumes instead of redoing the whole download
# (LibreOffice is hundreds of packages on a memory-constrained Docker VM).
FROM debian:bookworm-slim
ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update
RUN apt-get install -y --no-install-recommends fonts-liberation fonts-dejavu-core
RUN apt-get install -y --no-install-recommends libreoffice-core
RUN apt-get install -y --no-install-recommends libreoffice-writer libreoffice-calc \
 && apt-get clean \
 && rm -rf /var/lib/apt/lists/*

# Run as a non-root, no-login user.
RUN useradd --create-home --uid 1000 --shell /usr/sbin/nologin sandbox
USER sandbox
ENV HOME=/home/sandbox
WORKDIR /work

# Convert /in/<file> → /out/<file>.pdf. The LibreOffice user profile lives in a
# tmpfs (/tmp), so the rootfs can stay read-only.
ENTRYPOINT ["soffice", "--headless", "--nologo", "--nofirststartwizard", \
            "-env:UserInstallation=file:///tmp/loprofile", \
            "--convert-to", "pdf:writer_pdf_Export", "--outdir", "/out"]
