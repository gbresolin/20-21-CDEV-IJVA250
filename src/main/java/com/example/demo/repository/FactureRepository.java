package com.example.demo.repository;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * Repository permettant l'interraction avec la base de données pour les factures.
 */
@Repository
public interface FactureRepository extends JpaRepository<Facture, Long> {
    List<Facture> findAllById(Long id);
    List<Facture> countById(Long id);

}

