import argparse
import logging

from .config import get_settings
from .internet_mail import check_connections, run_self_test
from .workflow import (
    convert_response,
    prepare_request,
    process_input_messages,
    receive_responses,
    send_request,
)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    parser = argparse.ArgumentParser(description="Shipment forecast mail agent")
    subparsers = parser.add_subparsers(dest="command", required=True)

    prepare = subparsers.add_parser(
        "prepare", help="Convert input Excel to request JSON"
    )
    prepare.add_argument("input_xlsx")
    prepare.add_argument("request_json")
    prepare.add_argument("--reply-to", required=True)

    send = subparsers.add_parser(
        "send", help="Create JSON and send it from LogRocket to corporate Outlook"
    )
    send.add_argument("input_xlsx")

    subparsers.add_parser(
        "process-inbox",
        help="Read input XLSX mail from LogRocket and forward JSON to Outlook",
    )

    subparsers.add_parser(
        "receive", help="Read n8n responses from LogRocket and create Excel results"
    )

    subparsers.add_parser(
        "check-mail",
        help="Authenticate to IMAP and SMTP without sending a message",
    )
    subparsers.add_parser(
        "self-test-mail",
        help="Send a JSON message to the mailbox and read it back over IMAP",
    )

    convert = subparsers.add_parser(
        "convert", help="Convert a local response JSON to Excel"
    )
    convert.add_argument("response_json")
    convert.add_argument("output_xlsx")

    args = parser.parse_args()
    if args.command == "prepare":
        request = prepare_request(args.input_xlsx, args.request_json, args.reply_to)
        print(request.request_id)
    elif args.command == "send":
        request = send_request(args.input_xlsx, get_settings())
        print(request.request_id)
    elif args.command == "process-inbox":
        for request in process_input_messages(get_settings()):
            print(request.request_id)
    elif args.command == "receive":
        for path in receive_responses(get_settings()):
            print(path)
    elif args.command == "convert":
        print(convert_response(args.response_json, args.output_xlsx))
    elif args.command == "check-mail":
        settings = get_settings()
        for label, mailbox in (
            ("Exchange with n8n", settings),
            ("Input and results", settings.input_mailbox()),
        ):
            result = check_connections(mailbox)
            print(
                f"{label} ({mailbox.MAILBOX_NAME}): IMAP: {result['imap']}; SMTP: {result['smtp']}; authentication only, not delivery"
            )
    elif args.command == "self-test-mail":
        result = run_self_test(get_settings())
        print(
            f"SMTP: {result['smtp']}; IMAP: {result['imap']}; "
            f"JSON: {result['json']}; test_id: {result['test_id']}"
        )


if __name__ == "__main__":
    main()
