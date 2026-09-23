package com.example.service;

public class TransactionProcessor {

    public enum AuthResult {
        APPROVED,
        DENIED_INVALID_TOKEN,
        DENIED_INSUFFICIENT_FUNDS,
        DENIED_SYSTEM_OFFLINE
    }

    /**
     * Validates and processes a transaction request.
     */
    public AuthResult processTransaction(
            String authToken, 
            double amount, 
            double currentBalance, 
            boolean isServiceOnline) {

        if (authToken == null || authToken.trim().isEmpty()) {
            return AuthResult.DENIED_INVALID_TOKEN;
        }

        if (!isServiceOnline) {
            return AuthResult.DENIED_SYSTEM_OFFLINE;
        }

        if (amount <= 0 || amount > currentBalance) {
            return AuthResult.DENIED_INSUFFICIENT_FUNDS;
        }

        return AuthResult.APPROVED;
    }

    /**
     * Calculates the updated account balance.
     */
    public double calculateNewBalance(double balance, double deduction) {
        if (deduction < 0) {
            throw new IllegalArgumentException("Deduction amount cannot be negative.");
        }
        if (deduction > balance) {
            throw new IllegalStateException("Deduction cannot exceed available balance.");
        }
        return balance - deduction;
    }
}
