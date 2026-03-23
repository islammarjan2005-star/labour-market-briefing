FROM 165562107270.dkr.ecr.eu-west-2.amazonaws.com/jupyterhub-visualisation-base:rv4

ENV DATABASE_DSN__datasets_1=$DATABASE_DSN__datasets_1

WORKDIR /app

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

COPY . /app

CMD ["R", "-e", "shiny::runApp('/app', host = '0.0.0.0', port = 8888)"]

USER dw-user
