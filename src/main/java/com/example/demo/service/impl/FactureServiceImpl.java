package com.example.demo.service.impl;

import com.example.demo.entity.Facture;
import com.example.demo.repository.FactureRepository;
import com.example.demo.service.FactureService;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

/**
 * Service contenant les actions métiers liées aux factures.
 */
@Service
@Transactional
public class FactureServiceImpl implements FactureService {

    private FactureRepository factureRepository;

    public FactureServiceImpl(FactureRepository factureRepository) {
        this.factureRepository = factureRepository;
    }

    @Override
    public List<Facture> findAllFactures() {
        return factureRepository.findAll();
    }

    @Override
    public List<Facture> findFacturesNom() {
        return factureRepository.findAll();
    }
}
