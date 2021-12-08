SELECT *
FROM ((actor a
INNER JOIN film_actor fa on fa.actor_id=a.actor_id)
INNER JOIN film f on f.film_id=fa.film_id)
;
