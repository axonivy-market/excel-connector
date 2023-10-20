# Excel importer

Imports Excel sheets and transforms it into a full featured web application.

## Demo

TBD

## Setup

In the project, where the Excel data should be managed:

1. Create a persistence unit under `/config/persistence.xml`
2. Add the property, to allow schema changes `hibernate.hbm2ddl.auto=create`
3. Set the Data source to a valid database. If there is none, set it up under `/config/databases.yaml`