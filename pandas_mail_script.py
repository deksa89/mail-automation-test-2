import os
import smtplib
import ssl
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv


def load_env():
    load_dotenv()
    cfg = {
        "host": os.getenv("SMTP_HOST", "smtp.gmail.com"),
        "port": int(os.getenv("SMTP_PORT", "587")),
        "user": os.getenv("SMTP_USER"),
        "password": os.getenv("SMTP_PASS"),
        "from_email": os.getenv("FROM_EMAIL"),
        "fallback_to": os.getenv("FALLBACK_TO"),
    }
    missing = [k for k, v in cfg.items() if v in (None, "") and k not in ("port",)]
    if missing:
        raise RuntimeError(f"Missing required environment variables: {', '.join(missing)}")
    return cfg


def load_recipients(csv_path: Path) -> pd.DataFrame:
    """
    Reads recipients.csv and returns a DataFrame of active recipients with valid emails.
    If the file doesn't exist or yields no valid recipients, returns empty DataFrame.
    """
    if not csv_path.exists():
        return pd.DataFrame(columns=["email", "name"])

    df = pd.read_csv(csv_path, dtype={"email": "string", "name": "string"})
    # Normalize and filter
    df["email"] = df["email"].str.strip().str.lower()
    if "active" in df.columns:
        df = df[df["active"] == True]  # noqa: E712 (keep explicit True)
    # Basic email sanity
    df = df[df["email"].str.contains("@", na=False)]
    # Drop dupes by email, keep first
    df = df.drop_duplicates(subset=["email"])
    # Fill name (optional)
    if "name" not in df.columns:
        df["name"] = ""
    df["name"] = df["name"].fillna("")
    return df[["email", "name"]]


def build_report() -> pd.DataFrame:
    """
    Replace this with your real business logic.
    Here we do a tiny example that could be e.g. stock, orders, etc.
    """
    data = [
        {"sku": "D-235700", "status": "in_stock", "qty": 12},
        {"sku": "D-226994", "status": "low_stock", "qty": 2},
        {"sku": "D-211744", "status": "out_of_stock", "qty": 0},
    ]
    df = pd.DataFrame(data)
    # Example pandas processing: summarize counts by status
    summary = df.groupby("status", as_index=False).agg(total_skus=("sku", "count"), total_qty=("qty", "sum"))
    # Return both raw and summary concatenated for convenience
    summary.insert(0, "section", "summary")
    df2 = df.copy()
    df2.insert(0, "section", "detail")
    return pd.concat([summary, df2], ignore_index=True)


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def make_email(
    from_email: str,
    to_email: str,
    subject: str,
    plain_body: str,
    html_body: str | None = None,
    attachments: list[tuple[str, bytes]] | None = None,
) -> MIMEMultipart:
    msg = MIMEMultipart("mixed")
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject

    # Alternative part for plain + HTML
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(plain_body, "plain", "utf-8"))
    if html_body:
        alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt)

    for filename, content in attachments or []:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(content)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    return msg


def send_email(cfg: dict, message: MIMEMultipart):
    context = ssl.create_default_context()
    with smtplib.SMTP(cfg["host"], cfg["port"]) as server:
        server.starttls(context=context)
        server.login(cfg["user"], cfg["password"])
        server.send_message(message)


def main():
    cfg = load_env()

    # --- 1) Use pandas to load & process recipients
    recipients_df = load_recipients(Path("recipients.csv"))

    # Fallback if no recipients.csv or no active emails
    if recipients_df.empty:
        recipients_df = pd.DataFrame(
            [{"email": cfg["fallback_to"], "name": "Primatelj"}]
        )

    # --- 2) Use pandas to build a small report and attach it
    report_df = build_report()
    report_bytes = dataframe_to_csv_bytes(report_df)

    # Optional: build a tiny HTML preview table (first 10 rows)
    preview_html = report_df.head(10).to_html(index=False, justify="left")

    # --- 3) Send personalized emails (one per recipient)
    for _, row in recipients_df.iterrows():
        email = row["email"]
        name = (row.get("name") or "").strip()
        greeting_name = name if name else "pozdrav"

        # Your message text (customize freely)
        plain = (
            f"Bok {greeting_name},\n\n"
            "Bok kak si Deane? Ovo je test.\n\n"
            "U privitku je kratki report koji je generiran pomoću pandas-a.\n"
            "Lp!"
        )
        html = f"""
        <html>
          <body>
            <p>Bok {greeting_name},</p>
            <p><strong>Bok kak si Deane? Ovo je test.</strong></p>
            <p>U privitku je kratki report (CSV) generiran pomoću <code>pandas</code>.</p>
            <h4>Pregled (prvih 10 redaka)</h4>
            {preview_html}
            <p>LP!</p>
          </body>
        </html>
        """

        msg = make_email(
            from_email=cfg["from_email"],
            to_email=email,
            subject="Test s pandas izvještajem",
            plain_body=plain,
            html_body=html,
            attachments=[("report.csv", report_bytes)],
        )

        try:
            send_email(cfg, msg)
            print(f"✅ Sent to {email}")
        except Exception as e:
            print(f"❌ Failed to send to {email}: {e}")


if __name__ == "__main__":
    main()
