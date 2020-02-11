package com.example.demo.repository;

import com.example.demo.entity.Client;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * Repository permettant l'interraction avec la base de données pour les clients.
 */
@Repository
public interface ClientRepository extends JpaRepository<Client, Long> {
    List<Client> findAllById(Long id);
    List<Client> findByNomIgnoreCase(String nom);
}
