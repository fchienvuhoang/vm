This is a [Next.js](https://nextjs.org) project bootstrapped with [`create-next-app`](https://nextjs.org/docs/app/api-reference/cli/create-next-app).

## Getting Started

First, run the development server:

```bash
npm run dev
# or
yarn dev
# or
pnpm dev
# or
bun dev
```

Open [http://localhost:3000](http://localhost:3000) with your browser to see the result.

## Google Drive / Sheets integration

1. Enable the Google Drive API and Google Sheets API in a Google Cloud project.
2. Create a service account and a JSON key.
3. Share the destination Drive folder with the service account email as **Editor**.
4. Copy `.env.example` to `.env.local` and fill in `GOOGLE_SERVICE_ACCOUNT_EMAIL`, `GOOGLE_PRIVATE_KEY`, and `GOOGLE_DRIVE_FOLDER_ID`.
5. Restart the development server. Each category can then be linked to one spreadsheet and one tab.

The destination tab must have these headers in row 1 and in this exact order:

`STT | Thời gian | Chủ tài khoản | Nội dung | Tiền ra | Tiền vào | Ghi chú | Quỹ`

The app creates a hidden `_VIMUTTI_SYNC` tab in each destination spreadsheet. It stores transaction fingerprints used to skip duplicates without changing the required headers in the destination tab.

When syncing, the source transaction reference is appended on a new line at the end of the `Nội dung` cell (for example, `Số tham chiếu: FT...`). The destination schema therefore does not need another column.

You can start editing the page by modifying `app/page.tsx`. The page auto-updates as you edit the file.

This project uses [`next/font`](https://nextjs.org/docs/app/building-your-application/optimizing/fonts) to automatically optimize and load [Geist](https://vercel.com/font), a new font family for Vercel.

## Learn More

To learn more about Next.js, take a look at the following resources:

- [Next.js Documentation](https://nextjs.org/docs) - learn about Next.js features and API.
- [Learn Next.js](https://nextjs.org/learn) - an interactive Next.js tutorial.

You can check out [the Next.js GitHub repository](https://github.com/vercel/next.js) - your feedback and contributions are welcome!

## Deploy on Vercel

The easiest way to deploy your Next.js app is to use the [Vercel Platform](https://vercel.com/new?utm_medium=default-template&filter=next.js&utm_source=create-next-app&utm_campaign=create-next-app-readme) from the creators of Next.js.

Check out our [Next.js deployment documentation](https://nextjs.org/docs/app/building-your-application/deploying) for more details.
