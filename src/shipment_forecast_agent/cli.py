import argparse
import logging

from .agents import process_requests, process_results, run_agent
from .config import get_settings
from .internet_mail import check_connections, run_self_test
from .service import run_service
from .workflow import (
    convert_response,
    prepare_request,
    send_request,
)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    parser = argparse.ArgumentParser(description="Shipment forecast mail agent")
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser(
        "run-agents", help="Run both mail agents continuously; Ctrl+C stops both"
    )
    for name in ("request-agent", "result-agent"):
        purpose = (
            "prepare one incoming Excel request"
            if name == "request-agent"
            else "send one queued JSON or finish one response"
        )
        agent = subparsers.add_parser(
            name, help=f"Run once to {purpose}; use --continuous to keep polling"
        )
        mode = agent.add_mutually_exclusive_group()
        mode.add_argument(
            "--continuous",
            action="store_true",
            help="Keep polling until Ctrl+C instead of exiting after one cycle",
        )
        mode.add_argument(
            "--once",
            action="store_true",
            help=argparse.SUPPRESS,
        )

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
        help="Read one input email and queue its XLSX as JSON",
    )

    subparsers.add_parser(
        "receive", help="Send one queued JSON or finish one n8n response"
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
    if args.command == "run-agents":
        run_service(get_settings())
    elif args.command in ("request-agent", "result-agent"):
        try:
            run_agent(
                get_settings(),
                "requests" if args.command == "request-agent" else "results",
                once=not args.continuous,
            )
        except KeyboardInterrupt:
            print("Agent stopped")
    elif args.command == "prepare":
        request = prepare_request(args.input_xlsx, args.request_json, args.reply_to)
        print(request.request_id)
    elif args.command == "send":
        request = send_request(args.input_xlsx, get_settings())
        print(request.request_id)
    elif args.command == "process-inbox":
        for request_id in process_requests(get_settings()):
            print(request_id)
    elif args.command == "receive":
        for path in process_results(get_settings()):
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
