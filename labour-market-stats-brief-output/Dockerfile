FROM 165562107270.dkr.ecr.eu-west-2.amazonaws.com/jupyterhub-visualisation-base:rv4

ENV DATABASE_DSN__datasets_1=$DATABASE_DSN__datasets_1

WORKDIR /app

# Install R packages
RUN Rscript -e 'options(warn = 2); install.packages(c( \
  "shiny", \
  "DBI", \
  "RPostgres", \
  "dplyr", \
  "lubridate", \
  "glue", \
  "tibble", \
  "scales", \
  "readxl", \
  "openxlsx", \
  "officer", \
  "xml2" \
), clean = TRUE)'

# Install Python and openpyxl for perfect Excel copying
RUN apt-get update && apt-get install -y python3 python3-pip --no-install-recommends \
  && pip3 install openpyxl --break-system-packages \
  && apt-get clean \
  && rm -rf /var/lib/apt/lists/*

COPY . /app

CMD ["R", "-e", "shiny::runApp('/app', host = '0.0.0.0', port = 8888)"]

USER dw-user
