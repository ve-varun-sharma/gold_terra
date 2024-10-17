from email_validator import validate_email, EmailNotValidError, EmailUndeliverableError


def email_validator(email: str) -> bool:
    try:

        # Check that the email address is valid. Turn on check_deliverability
        # for first-time validations like on account creation pages (but not
        # login pages).
        emailinfo = validate_email(email, check_deliverability=True)

        # After this point, use only the normalized form of the email address,
        # especially before going to a database query.
        email = emailinfo.normalized
        print(emailinfo)
        print(email)

        return True

    except EmailUndeliverableError as undeliverable_error:
        print(undeliverable_error)
        return False

    except EmailNotValidError as e:

        # The exception message is human-readable explanation of why it's
        # not a valid (or deliverable) email address.
        print(str(e))
        return False

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return False


email = "plowboy1944@gmail.com"

email_validator(email)
