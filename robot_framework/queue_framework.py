"""This module is the primary module of the robot framework. It collects the functionality of the rest of the framework."""

# This module is not meant to exist next to linear_framework.py in production:
# pylint: disable=duplicate-code

import sys

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus

from robot_framework import initialize
from robot_framework import reset
from robot_framework.exceptions import handle_error, BusinessError, log_exception
from robot_framework import process
from robot_framework import config
import json
import smtplib
from email.message import EmailMessage
import os

def main():
    """The entry point for the framework. Should be called as the first thing when running the robot."""
    orchestrator_connection = OrchestratorConnection.create_connection_from_args()
    sys.excepthook = log_exception(orchestrator_connection)

    orchestrator_connection.log_trace("Robot Framework started.")
    initialize.initialize(orchestrator_connection)

    queue_element = None
    error_count = 0
    task_count = 0
    # Retry loop
    for _ in range(config.MAX_RETRY_COUNT):
        try:
            reset.reset(orchestrator_connection)

            # Queue loop
            while task_count < config.MAX_TASK_COUNT:
                task_count += 1
                queue_element = orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)

                if not queue_element:
                    orchestrator_connection.log_info("Queue empty.")
                    break  # Break queue loop

                try:
                    for attempt in range(1, config.QUEUE_ATTEMPTS + 1):
                        try:
                            process.process(orchestrator_connection, queue_element)
                            break
                        except Exception as e:
                                #Deleting potential leftover files from downloads folder
                            orchestrator_connection.log_info('Deleting local files')

                            downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
                            specific_content = json.loads(queue_element.data)
                            UdviklerMail = orchestrator_connection.get_constant("Error Email").value
                            MailModtager = specific_content.get("Ansvarlig i Økonomi", None)
                            MailModtager = UdviklerMail
                            FileName = specific_content.get("Filnavn", None)
                            if os.path.exists(downloads_folder + '\\' + FileName + ".xls"):
                                os.remove(downloads_folder + '\\' + FileName + ".xls")
                            if os.path.exists(downloads_folder + '\\' + "YKMD_STD.xls"):
                                os.remove(downloads_folder + '\\' + "YKMD_STD.xls")
                            orchestrator_connection.log_trace(f"Attempt {attempt} failed for current queue element: {e}")
                            if attempt < config.QUEUE_ATTEMPTS:
                                orchestrator_connection.log_trace("Retrying queue element.")
                                reset.reset(orchestrator_connection)
                            else:
                                orchestrator_connection.log_trace(f"Queue element failed after {attempt} attempts.")
                                
                                send_error_email(MailModtager, FileName, UdviklerMail)
                                raise
                    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.DONE)

                except BusinessError as error:
                    handle_error("Business Error", error, queue_element, orchestrator_connection)

            break  # Break retry loop

        # We actually want to catch all exceptions possible here.
        # pylint: disable-next = broad-exception-caught
        except Exception as error:
            error_count += 1
            handle_error(f"Process Error #{error_count}", error, queue_element, orchestrator_connection)

    reset.clean_up(orchestrator_connection)
    reset.close_all(orchestrator_connection)
    reset.kill_all(orchestrator_connection)

    if config.FAIL_ROBOT_ON_TOO_MANY_ERRORS and error_count == config.MAX_RETRY_COUNT:
        raise RuntimeError("Process failed too many times.")
    
def send_error_email(to_address: str | list[str], file_name: str, UdviklerMail ):
    """
    Sends an email notification with the provided body and subject.

    Args:
        to_address (str | list[str]): Email address or list of addresses to send the notification.
        sags_id (str): The ID of the case (SagsID) used in the email subject.
        deskpro_id (str): The DeskPro ID for constructing the DeskPro link.
        sharepoint_link (str): The SharePoint link to include in the email body.
    """
    # SMTP Configuration (from your provided details)
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "opus@aarhus.dk"
    # Email subject
    subject = f"Fejl i processeringen af filen {file_name}"

    # Email body (HTML)
    body = f"""
    <html>
    <body>
        <p>Der var en fejl i processeringen af filen {file_name}. Hvis fejlen fortsætter, kontakt udvikler</p>
    </body>
    </html>
    """

    # Create the email message
    msg = EmailMessage()
    msg['To'] = ', '.join(to_address) if isinstance(to_address, list) else to_address
    msg['From'] = SCREENSHOT_SENDER
    msg['Subject'] = subject
    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(body, subtype='html')
    msg['Reply-To'] = UdviklerMail
    msg['Bcc'] = UdviklerMail

    # Send the email using SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.send_message(msg)
            
    except Exception as e:
        print(f"Failed to send success email: {e}")