export interface ParsedOrder {
  id: number;
  created_at: string;
  company_name: string;
  contact_name: string;
  contact_email: string;
  contact_phone?: string;
  description: string;
  budget_from?: number;
  budget_to?: number;
  deadline?: string;
  raw_email?: string;
}
