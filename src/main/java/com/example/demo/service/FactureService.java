package com.example.demo.service;

import com.example.demo.entity.Facture;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface FactureService {
    List<Facture> findAllFactures();

    @Query("SELECT * FROM FACTURE JOIN CLIENT ON client.id = facture.client_ID")
    List<Facture> findFacturesNom();
}
