package com.example.service;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class TransactionProcessorTest {

    private TransactionProcessor processor;

    @BeforeEach
    void setUp() {
        processor = new TransactionProcessor();
    }

    @Test
    @DisplayName("REQ-001: Valid token and sufficient balance is approved")
    void testValidTransactionApproved() {
        TransactionProcessor.AuthResult result = 
            processor.processTransaction("TOKEN_VALID_123", 50.0, 100.0, true);

        assertEquals(TransactionProcessor.AuthResult.APPROVED, result);
    }

    @Test
    @DisplayName("REQ-002: Reject when authentication token is empty or null")
    void testInvalidTokenRejected() {
        TransactionProcessor.AuthResult result = 
            processor.processTransaction("", 50.0, 100.0, true);

        assertEquals(TransactionProcessor.AuthResult.DENIED_INVALID_TOKEN, result);
    }

    @Test
    @DisplayName("REQ-003: Reject when backend system is offline")
    void testSystemOfflineRejected() {
        TransactionProcessor.AuthResult result = 
            processor.processTransaction("TOKEN_VALID_123", 50.0, 100.0, false);

        assertEquals(TransactionProcessor.AuthResult.DENIED_SYSTEM_OFFLINE, result);
    }

    @Test
    @DisplayName("REQ-004: Reject when amount exceeds available balance")
    void testInsufficientFundsRejected() {
        TransactionProcessor.AuthResult result = 
            processor.processTransaction("TOKEN_VALID_123", 150.0, 100.0, true);

        assertEquals(TransactionProcessor.AuthResult.DENIED_INSUFFICIENT_FUNDS, result);
    }

    @Test
    @DisplayName("REQ-005: Correctly deduct balance on transaction")
    void testBalanceDeduction() {
        double updated = processor.calculateNewBalance(200.0, 75.0);
        assertEquals(125.0, updated, 0.001);
    }

    @Test
    @DisplayName("REQ-006: Throw exception when deduction exceeds balance")
    void testOverdraftThrowsException() {
        assertThrows(IllegalStateException.class, () -> {
            processor.calculateNewBalance(50.0, 100.0);
        });
    }
}
