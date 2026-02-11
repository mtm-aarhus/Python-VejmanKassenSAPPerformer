# README — SAP Invoice Processing Flow

## Overview

This robot processes creates invoices from the dispatcher. For each queue item, it creates an invoice csv file, uploads it to SAP, and if there's no errors it sends it and updates it in VejmanKassen. If SAP rejects the invoice due to missing debitor, the robot attempts to create the required debitors and retries. If it still cannot complete successfully, it raises an error and stops so the case can be handled manually.

## High-level flow

1. **Startup**

   * Initializes SAP (ensures a usable SAP GUI session is available).

2. **Queue processing**

   * Fetches the next queue element until the queue is empty or a maximum task limit is reached.
   * For each queue element:

     * Loads invoice identifiers from the queue payload (e.g., SQL row ID and case reference).
     * Looks up the invoice row in the database (only rows in status **TilFakturering** are processed).
     * Generates an invoice CSV (invoice header + line(s)) for SAP import.

3. **Create invoice in SAP**

   * Runs the SAP transaction to import the generated invoice file.
   * **If the file is accepted:**

     * SAP creates exactly one order number (validated).
   * **If the file is not accepted due to missing debitor:**

     * Extracts the affected CVR/debitor identifiers from SAP’s error list.
     * Generates a debitor CSV and runs the SAP debitor creation transaction.
     * Retries the invoice import once again after debitors are created.
   * **If the invoice still cannot be created:**

     * Raises an error and stops processing for manual handling.

4. **Send invoice**

   * Executes the SAP sending step (runs the relevant send/commit flow).
   * Validates there are no SAP “Fejl” entries before completing.

5. **Finalize**

   * Updates the database row in VejmanKassen:

     * Sets status to **Faktureret**
     * Stores invoice date
     * Stores SAP order number
   * Removes the temporary invoice CSV file.
   * If the case is a Vejman case, it updates the case to reflect that the invoice was sent.


## Outputs and side effects

* SAP invoice creation and sending.
* Database updates for processed invoices.
* Optional case update to reflect “invoice sent”.

# Robot-Framework V4

This repo is meant to be used as a template for robots made for [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator) v2.

## Quick start

1. To use this template simply use this repo as a template (see [Creating a repository from a template](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-repository-from-a-template)).
__Don't__ include all branches.

2. Go to `robot_framework/__main__.py` and choose between the linear framework or queue based framework.

3. Implement all functions in the files:
    * `robot_framework/initialize.py`
    * `robot_framework/reset.py`
    * `robot_framework/process.py`

4. Change `config.py` to your needs.

5. Fill out the dependencies in the `pyproject.toml` file with all packages needed by the robot.

6. Feel free to add more files as needed. Remember that any additional python files must
be located in the folder `robot_framework` or a subfolder of it.

When the robot is run from OpenOrchestrator the `main.py` file is run which results
in the following:

1. The working directory is changed to where `main.py` is located.
2. A virtual environment is automatically setup with the required packages.
3. The framework is called passing on all arguments needed by [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

## Requirements

Minimum python version 3.11

## Flow

This framework contains two different flows: A linear and a queue based.
You should only ever use one at a time. You choose which one by going into `robot_framework/__main__.py`
and uncommenting the framework you want. They are both disabled by default and an error will be
raised to remind you if you don't choose.

### Linear Flow

The linear framework is used when a robot is just going from A to Z without fetching jobs from an
OpenOrchestrator queue.
The flow of the linear framework is sketched up in the following illustration:

![Linear Flow diagram](Robot-Framework.svg)

### Queue Flow

The queue framework is used when the robot is doing multiple bite-sized tasks defined in an
OpenOrchestrator queue.
The flow of the queue framework is sketched up in the following illustration:

![Queue Flow diagram](Robot-Queue-Framework.svg)

## Linting and Github Actions

This template is also setup with flake8 and pylint linting in Github Actions.
This workflow will trigger whenever you push your code to Github.
The workflow is defined under `.github/workflows/Linting.yml`.
