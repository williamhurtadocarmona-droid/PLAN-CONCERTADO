if len(doc.tables) > 1:
                        t_act = doc.tables[1]
                        f_conc_str = fecha_concertada.strftime("%d/%m/%Y")
                        f_fin_str = fecha_final.strftime("%d/%m/%Y")

                        for idx, desc in enumerate(acts_desc):
                            r_idx = idx + 3
                            if r_idx < len(t_act.rows):
                                cells = t_act.rows[r_idx].cells
                                cells[0].text = rap_sel
                                cells[1].text = str(idx + 1)
                                cells[2].text = desc
                                
                                # Forma de entrega (Física / Digital)
                                cells[3].text = "X" if ent_act[idx] == "Física" else ""
                                cells[4].text = "" if ent_act[idx] == "Física" else "X"
                                
                                # Fechas de entrega (Concertada y Final) en las columnas correspondientes
                                cells[5].text = f_conc_str
                                cells[6].text = f_fin_str

                                # ¿Entregó la Actividad? (SI / NO)
                                if tipo_plan == "Plan Final":
                                    cells[7].text = "X" if est_act[idx] == "SÍ" else ""
                                    cells[8].text = "X" if est_act[idx] == "NO" else ""
                                else:
                                    cells[7].text = ""
                                    cells[8].text = ""

                        for r_rm in range(12, len(acts_desc) + 2, -1):
                            if r_rm < len(t_act.rows):
                                tr = t_act.rows[r_rm]._tr
                                tr.getparent().remove(tr)
