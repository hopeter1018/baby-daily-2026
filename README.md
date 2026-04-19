# baby-daily-2026

A baby-caring daily app written with **React + Tailwind**, powered by **Firebase Auth** and **Cloud Firestore**.

## Features

- Google sign-in via Firebase Auth
- Daily baby care logs stored in Firestore
- `schemaVersion` on every record for data-structure evolution
- XLSX report export
- XLSX email flow to one specific email (`VITE_REPORT_EMAIL`)

## Setup

1. Install dependencies:

   ```bash
   npm install
   ```

2. Add environment variables (copy from `.env.example`):

   ```bash
   cp .env.example .env
   ```

3. Start development server:

   ```bash
   npm run dev
   ```

## Commands

- `npm run lint`
- `npm run build`
- `npm run preview`
